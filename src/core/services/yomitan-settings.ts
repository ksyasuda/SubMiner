import electron from 'electron';
import type { BrowserWindow, Extension, Menu, MenuItemConstructorOptions, Session } from 'electron';
import { createLogger } from '../../logger';
import { promoteSettingsWindowAboveOverlay } from './settings-window-z-order';

const { BrowserWindow: ElectronBrowserWindow, Menu: ElectronMenu, session } = electron;
const logger = createLogger('main:yomitan-settings');

export interface OpenYomitanSettingsWindowOptions {
  yomitanExt: Extension | null;
  getExistingWindow: () => BrowserWindow | null;
  setWindow: (window: BrowserWindow | null) => void;
  yomitanSession?: Session | null;
  onWindowClosed?: () => void;
}

type YomitanSettingsWindowMenuOwner = Pick<BrowserWindow, 'close' | 'isDestroyed'>;

type HyprlandSessionEnv = {
  HYPRLAND_INSTANCE_SIGNATURE?: string;
};

export interface InstallYomitanSettingsCloseButtonOptions {
  platform?: NodeJS.Platform;
  env?: HyprlandSessionEnv;
}

export function shouldInstallYomitanSettingsCloseButton(
  platform: NodeJS.Platform = process.platform,
  env: HyprlandSessionEnv = process.env,
): boolean {
  return platform === 'linux' && Boolean(env.HYPRLAND_INSTANCE_SIGNATURE);
}

export function buildYomitanSettingsWindowMenuTemplate(
  settingsWindow: YomitanSettingsWindowMenuOwner,
): MenuItemConstructorOptions[] {
  return [
    {
      label: 'File',
      submenu: [
        {
          label: 'Close',
          accelerator: process.platform === 'darwin' ? 'Command+W' : 'Ctrl+W',
          click: () => {
            if (!settingsWindow.isDestroyed()) {
              settingsWindow.close();
            }
          },
        },
      ],
    },
  ];
}

export function buildYomitanSettingsCloseButtonScript(): string {
  return `
(() => {
  const buttonId = 'subminer-yomitan-settings-close';
  const styleId = 'subminer-yomitan-settings-close-style';
  if (document.getElementById(buttonId)) {
    return;
  }
  if (!document.getElementById(styleId)) {
    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = \`
      #\${buttonId} {
        position: fixed;
        top: 10px;
        left: 10px;
        z-index: 2147483647;
        width: 32px;
        height: 32px;
        display: grid;
        place-items: center;
        padding: 0;
        border: 1px solid rgba(255, 255, 255, 0.28);
        border-radius: 4px;
        background: rgba(24, 24, 24, 0.92);
        color: #f2f2f2;
        font: 22px/1 system-ui, sans-serif;
        cursor: pointer;
      }
      #\${buttonId}:hover {
        background: rgba(54, 54, 54, 0.96);
        border-color: rgba(255, 255, 255, 0.5);
      }
      #\${buttonId}:focus-visible {
        outline: 2px solid #8ab4f8;
        outline-offset: 2px;
      }
    \`;
    document.head.appendChild(style);
  }
  const button = document.createElement('button');
  button.id = buttonId;
  button.type = 'button';
  button.title = 'Close';
  button.setAttribute('aria-label', 'Close Yomitan settings');
  button.textContent = '\\u00d7';
  button.addEventListener('click', () => {
    window.close();
  });
  document.body.appendChild(button);
})();
`;
}

export function installYomitanSettingsCloseButton(
  settingsWindow: Pick<BrowserWindow, 'isDestroyed' | 'webContents'>,
  options: InstallYomitanSettingsCloseButtonOptions = {},
): void {
  if (settingsWindow.isDestroyed()) {
    return;
  }
  if (!shouldInstallYomitanSettingsCloseButton(options.platform, options.env)) {
    return;
  }
  settingsWindow.webContents
    .executeJavaScript(buildYomitanSettingsCloseButtonScript())
    .catch((error: Error) => {
      logger.warn('Failed to install Yomitan settings close button:', error.message);
    });
}

export function configureYomitanSettingsWindowChrome(
  settingsWindow: Pick<BrowserWindow, 'close' | 'isDestroyed' | 'setAutoHideMenuBar' | 'setMenu'>,
  buildMenu: (template: MenuItemConstructorOptions[]) => Menu = (template) =>
    ElectronMenu.buildFromTemplate(template),
): void {
  settingsWindow.setAutoHideMenuBar(false);
  settingsWindow.setMenu(buildMenu(buildYomitanSettingsWindowMenuTemplate(settingsWindow)));
}

export function buildYomitanSettingsUrl(extensionId: string): string {
  return `chrome-extension://${extensionId}/settings.html?popup-preview=false`;
}

export function showYomitanSettingsWindow(
  settingsWindow: BrowserWindow,
  options: {
    promoteSettingsWindowAboveOverlay?: (settingsWindow: BrowserWindow) => void;
  } = {},
): void {
  if (settingsWindow.isDestroyed()) {
    return;
  }
  if (settingsWindow.isMinimized()) {
    settingsWindow.restore();
  }
  const [width = 0, height = 0] = settingsWindow.getSize();
  settingsWindow.setSize(width, height);
  settingsWindow.webContents.invalidate();
  settingsWindow.show();
  settingsWindow.focus();
  (options.promoteSettingsWindowAboveOverlay ?? promoteSettingsWindowAboveOverlay)(settingsWindow);
}

export function destroyYomitanSettingsWindow(settingsWindow: BrowserWindow | null): boolean {
  if (!settingsWindow || settingsWindow.isDestroyed()) {
    return false;
  }
  settingsWindow.destroy();
  return true;
}

export function openYomitanSettingsWindow(options: OpenYomitanSettingsWindowOptions): void {
  logger.info('openYomitanSettings called');

  if (!options.yomitanExt) {
    logger.error('Yomitan extension not loaded - yomitanExt is:', options.yomitanExt);
    logger.error('This may be due to Manifest V3 service worker issues with Electron');
    return;
  }

  const existingWindow = options.getExistingWindow();
  if (existingWindow && !existingWindow.isDestroyed()) {
    logger.info('Settings window already exists, showing and focusing');
    showYomitanSettingsWindow(existingWindow);
    return;
  }

  logger.info('Creating new settings window for extension:', options.yomitanExt.id);

  const settingsWindow = new ElectronBrowserWindow({
    title: 'Yomitan Settings',
    width: 1200,
    height: 800,
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      session: options.yomitanSession ?? session.defaultSession,
    },
  });
  configureYomitanSettingsWindowChrome(settingsWindow);
  options.setWindow(settingsWindow);

  const settingsUrl = buildYomitanSettingsUrl(options.yomitanExt.id);
  logger.info('Loading settings URL:', settingsUrl);

  let loadAttempts = 0;
  const maxAttempts = 3;

  const attemptLoad = (): void => {
    settingsWindow
      .loadURL(settingsUrl)
      .then(() => {
        logger.info('Settings URL loaded successfully');
      })
      .catch((err: Error) => {
        logger.error('Failed to load settings URL:', err);
        loadAttempts++;
        if (loadAttempts < maxAttempts && !settingsWindow.isDestroyed()) {
          logger.info(`Retrying in 500ms (attempt ${loadAttempts + 1}/${maxAttempts})`);
          setTimeout(attemptLoad, 500);
        }
      });
  };

  attemptLoad();

  settingsWindow.webContents.on('did-fail-load', (_event, errorCode, errorDescription) => {
    logger.error('Settings page failed to load:', errorCode, errorDescription);
  });

  settingsWindow.webContents.on('did-finish-load', () => {
    logger.info('Settings page loaded successfully');
    installYomitanSettingsCloseButton(settingsWindow);
  });

  setTimeout(() => {
    showYomitanSettingsWindow(settingsWindow);
  }, 500);

  settingsWindow.on('closed', () => {
    options.onWindowClosed?.();
    options.setWindow(null);
  });
}
