import type {
  BrowserWindow,
  BrowserWindowConstructorOptions,
  IpcMain,
  MenuItemConstructorOptions,
  Session,
} from 'electron';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  parseYoutubeBrowserVideoRequest,
  type YoutubeBrowserVideoRequest,
  type YoutubeBrowserVideoResult,
} from './youtube-browser-playback';
import { toYoutubeWatchUrl } from './youtube-playback';

/** Persistent partition: cookies (and so the YouTube login) survive app restarts. */
export const YOUTUBE_BROWSER_PARTITION = 'persist:youtube-browser';
export const YOUTUBE_BROWSER_HOME_URL = 'https://www.youtube.com/';

const ALLOWED_PERMISSIONS = new Set(['fullscreen', 'clipboard-sanitized-write']);
const STANDARD_UA_PRODUCTS = '(?:Mozilla|AppleWebKit|Chrome|Safari)';
const NON_BROWSER_UA_TOKEN = new RegExp(
  `\\s+(?!${STANDARD_UA_PRODUCTS}/)[A-Za-z][\\w.-]*/\\S+`,
  'g',
);

/**
 * Plain Chrome user agent. Google refuses sign-in from agents that advertise Electron or an
 * embedding app, so every product token other than the standard Chrome ones is dropped.
 */
export function buildYoutubeBrowserUserAgent(defaultUserAgent: string): string {
  return defaultUserAgent.replace(NON_BROWSER_UA_TOKEN, '').trim();
}

export type YoutubeBrowserWindowDeps = {
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => BrowserWindow;
  preloadPath: string;
  ipcMain: Pick<IpcMain, 'on'>;
  showContextMenu: (window: BrowserWindow, template: MenuItemConstructorOptions[]) => void;
  openExternal: (url: string) => void;
  openVideo: (request: YoutubeBrowserVideoRequest) => Promise<YoutubeBrowserVideoResult>;
  logWarn: (message: string, error?: unknown) => void;
  logDebug: (message: string) => void;
  onClosed?: () => void;
};

function isHttpUrl(url: string): boolean {
  return url.startsWith('https://') || url.startsWith('http://');
}

function configureSession(session: Session): void {
  session.setUserAgent(buildYoutubeBrowserUserAgent(session.getUserAgent()));
  session.setPermissionRequestHandler((_webContents, permission, callback) => {
    callback(ALLOWED_PERMISSIONS.has(permission));
  });
}

/**
 * A logged-in YouTube window whose video links open in mpv instead of the page player.
 * The preload intercepts link clicks; navigation hooks here catch anything that slips past it.
 */
export function createYoutubeBrowserWindowRuntime(deps: YoutubeBrowserWindowDeps) {
  let browserWindow: BrowserWindow | null = null;
  let ipcRegistered = false;

  // `trigger` names the path that caught the video, to tell duplicate requests apart in logs.
  const sendVideo = (request: YoutubeBrowserVideoRequest, trigger: string): void => {
    deps.logDebug(`YouTube browser ${request.action} via ${trigger}: ${request.url}`);
    void deps
      .openVideo(request)
      .catch((error: unknown): YoutubeBrowserVideoResult => {
        deps.logWarn('YouTube browser failed to open video in mpv', error);
        return { ok: false, message: 'Could not open the video in mpv.' };
      })
      .then((result) => {
        if (browserWindow && !browserWindow.isDestroyed()) {
          browserWindow.webContents.send(IPC_CHANNELS.event.youtubeBrowserToast, result);
        }
      });
  };

  const registerIpc = (): void => {
    if (ipcRegistered) return;
    ipcRegistered = true;
    deps.ipcMain.on(IPC_CHANNELS.command.youtubeBrowserOpenVideo, (event, payload: unknown) => {
      if (!browserWindow || event.sender !== browserWindow.webContents) return;
      const request = parseYoutubeBrowserVideoRequest(payload);
      if (request) sendVideo(request, 'click');
    });
  };

  const buildContextMenu = (window: BrowserWindow, linkUrl: string) => {
    const history = window.webContents.navigationHistory;
    const videoItems: MenuItemConstructorOptions[] = toYoutubeWatchUrl(linkUrl)
      ? [
          {
            label: 'Play in mpv',
            click: () => sendVideo({ action: 'play', url: linkUrl }, 'context-menu'),
          },
          {
            label: 'Queue in mpv',
            click: () => sendVideo({ action: 'queue', url: linkUrl }, 'context-menu'),
          },
          { type: 'separator' },
        ]
      : [];
    return [
      ...videoItems,
      { label: 'Back', enabled: history.canGoBack(), click: () => history.goBack() },
      { label: 'Forward', enabled: history.canGoForward(), click: () => history.goForward() },
      { label: 'Reload', click: () => window.webContents.reload() },
    ] satisfies MenuItemConstructorOptions[];
  };

  const createWindow = (): BrowserWindow => {
    const window = deps.createBrowserWindow({
      width: 1280,
      height: 860,
      title: 'SubMiner YouTube',
      autoHideMenuBar: true,
      backgroundColor: '#0f0f0f',
      webPreferences: {
        partition: YOUTUBE_BROWSER_PARTITION,
        preload: deps.preloadPath,
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: true,
      },
    });
    const { webContents } = window;
    const { session } = webContents;
    configureSession(session);

    webContents.setWindowOpenHandler(({ url }) => {
      if (toYoutubeWatchUrl(url)) {
        sendVideo({ action: 'queue', url }, 'window-open');
      } else if (isHttpUrl(url)) {
        deps.openExternal(url);
      }
      return { action: 'deny' };
    });

    // Full page loads of a video (e.g. from a redirect) never reach the preload's click hook.
    webContents.on('will-navigate', (event, url) => {
      if (!toYoutubeWatchUrl(url)) return;
      event.preventDefault();
      sendVideo({ action: 'play', url }, 'will-navigate');
    });

    // In-app navigations the preload missed (keyboard shortcuts, script-driven links): stop the
    // page player, hand the video to mpv, and return to the page the user came from.
    webContents.on('did-navigate-in-page', (_event, url, isMainFrame) => {
      if (!isMainFrame || !toYoutubeWatchUrl(url)) return;
      void webContents
        .executeJavaScript("document.querySelectorAll('video').forEach((v) => v.pause())")
        .catch(() => {});
      sendVideo({ action: 'play', url }, 'in-page-navigation');
      if (webContents.navigationHistory.canGoBack()) {
        webContents.navigationHistory.goBack();
      } else {
        void webContents.loadURL(YOUTUBE_BROWSER_HOME_URL).catch(() => {});
      }
    });

    webContents.on('context-menu', (_event, params) => {
      deps.showContextMenu(window, buildContextMenu(window, params.linkURL));
    });

    webContents.on('before-input-event', (event, input) => {
      if (input.type !== 'keyDown') return;
      const history = webContents.navigationHistory;
      if (input.alt && input.key === 'ArrowLeft' && history.canGoBack()) {
        event.preventDefault();
        history.goBack();
      } else if (input.alt && input.key === 'ArrowRight' && history.canGoForward()) {
        event.preventDefault();
        history.goForward();
      } else if (input.key === 'F5' || ((input.control || input.meta) && input.key === 'r')) {
        event.preventDefault();
        webContents.reload();
      }
    });

    window.on('closed', () => {
      browserWindow = null;
      void session.cookies.flushStore().catch((error: unknown) => {
        deps.logWarn('Failed to flush YouTube browser cookies', error);
      });
      deps.onClosed?.();
    });

    void window.loadURL(YOUTUBE_BROWSER_HOME_URL).catch((error: unknown) => {
      // YouTube redirects its landing page (e.g. `?themeRefresh=1`), which aborts the first load.
      if ((error as { code?: unknown }).code === 'ERR_ABORTED') return;
      deps.logWarn('Failed to load YouTube in the browser window', error);
    });
    return window;
  };

  const open = (): void => {
    registerIpc();
    if (browserWindow && !browserWindow.isDestroyed()) {
      if (browserWindow.isMinimized()) browserWindow.restore();
      browserWindow.show();
      browserWindow.focus();
      return;
    }
    browserWindow = createWindow();
  };

  const isOpen = (): boolean => Boolean(browserWindow && !browserWindow.isDestroyed());

  return { open, isOpen };
}
