import { BrowserWindow, screen, type Session } from 'electron';
import * as path from 'path';
import { WindowGeometry } from '../../types';
import { createLogger } from '../../logger';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  handleOverlayWindowBeforeInputEvent,
  handleOverlayWindowBlurred,
  type OverlayWindowKind,
} from './overlay-window-input';
import { buildOverlayWindowOptions } from './overlay-window-options';
import { normalizeOverlayWindowBoundsForPlatform } from './overlay-window-bounds';

const logger = createLogger('main:overlay-window');
const overlayWindowLayerByInstance = new WeakMap<BrowserWindow, OverlayWindowKind>();
const overlayWindowContentReady = new WeakSet<BrowserWindow>();
const OVERLAY_WINDOW_CONTENT_READY_FLAG = '__subminerOverlayContentReady';

export function isOverlayWindowContentReady(window: BrowserWindow): boolean {
  if (window.isDestroyed()) {
    return false;
  }
  return (
    overlayWindowContentReady.has(window) ||
    (window as BrowserWindow & { [OVERLAY_WINDOW_CONTENT_READY_FLAG]?: boolean })[
      OVERLAY_WINDOW_CONTENT_READY_FLAG
    ] === true
  );
}

function getOverlayWindowHtmlPath(): string {
  return path.join(__dirname, '..', '..', 'renderer', 'index.html');
}

function loadOverlayWindowLayer(window: BrowserWindow, layer: OverlayWindowKind): void {
  overlayWindowLayerByInstance.set(window, layer);
  const htmlPath = getOverlayWindowHtmlPath();
  window
    .loadFile(htmlPath, {
      query: { layer },
    })
    .catch((err) => {
      logger.error('Failed to load HTML file:', err);
    });
}

export function updateOverlayWindowBounds(
  geometry: WindowGeometry,
  window: BrowserWindow | null,
): void {
  if (!geometry || !window || window.isDestroyed()) return;
  window.setBounds(normalizeOverlayWindowBoundsForPlatform(geometry, process.platform, screen));
}

export function ensureOverlayWindowLevel(window: BrowserWindow): void {
  if (process.platform === 'darwin') {
    window.setAlwaysOnTop(true, 'screen-saver', 1);
    window.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
    window.setFullScreenable(false);
    window.moveTop();
    return;
  }
  if (process.platform === 'win32') {
    window.setAlwaysOnTop(true, 'screen-saver', 1);
    window.moveTop();
    return;
  }
  window.setAlwaysOnTop(true);
}

export function enforceOverlayLayerOrder(options: {
  visibleOverlayVisible: boolean;
  mainWindow: BrowserWindow | null;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
}): void {
  if (!options.visibleOverlayVisible) return;
  if (!options.mainWindow || options.mainWindow.isDestroyed()) return;

  options.ensureOverlayWindowLevel(options.mainWindow);
  options.mainWindow.moveTop();
}

export function createOverlayWindow(
  kind: OverlayWindowKind,
  options: {
    isDev: boolean;
    ensureOverlayWindowLevel: (window: BrowserWindow) => void;
    onRuntimeOptionsChanged: () => void;
    setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
    isOverlayVisible: (kind: OverlayWindowKind) => boolean;
    tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
    forwardTabToMpv: () => void;
    onVisibleWindowBlurred?: () => void;
    onWindowContentReady?: () => void;
    onWindowClosed: (kind: OverlayWindowKind) => void;
    yomitanSession?: Session | null;
  },
): BrowserWindow {
  const window = new BrowserWindow(buildOverlayWindowOptions(kind, options));

  if (!(process.platform === 'win32' && kind === 'visible')) {
    options.ensureOverlayWindowLevel(window);
  }
  loadOverlayWindowLayer(window, kind);

  window.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL) => {
    logger.error('Page failed to load:', errorCode, errorDescription, validatedURL);
  });

  window.webContents.on('did-finish-load', () => {
    options.onRuntimeOptionsChanged();
  });

  window.once('ready-to-show', () => {
    overlayWindowContentReady.add(window);
    (window as BrowserWindow & { [OVERLAY_WINDOW_CONTENT_READY_FLAG]?: boolean })[
      OVERLAY_WINDOW_CONTENT_READY_FLAG
    ] = true;
    options.onWindowContentReady?.();
  });

  if (kind === 'visible') {
    window.webContents.on('devtools-opened', () => {
      options.setOverlayDebugVisualizationEnabled(true);
    });
    window.webContents.on('devtools-closed', () => {
      options.setOverlayDebugVisualizationEnabled(false);
    });
  }

  window.webContents.on('before-input-event', (event, input) => {
    handleOverlayWindowBeforeInputEvent({
      kind,
      windowVisible: window.isVisible(),
      input,
      preventDefault: () => event.preventDefault(),
      sendKeyboardModeToggleRequested: () =>
        window.webContents.send(IPC_CHANNELS.event.keyboardModeToggleRequested),
      sendLookupWindowToggleRequested: () =>
        window.webContents.send(IPC_CHANNELS.event.lookupWindowToggleRequested),
      tryHandleOverlayShortcutLocalFallback: (nextInput) =>
        options.tryHandleOverlayShortcutLocalFallback(nextInput),
      forwardTabToMpv: () => options.forwardTabToMpv(),
    });
  });

  window.hide();

  window.on('closed', () => {
    options.onWindowClosed(kind);
  });

  window.on('blur', () => {
    if (window.isDestroyed()) return;
    handleOverlayWindowBlurred({
      kind,
      windowVisible: window.isVisible(),
      isOverlayVisible: options.isOverlayVisible,
      ensureOverlayWindowLevel: () => {
        options.ensureOverlayWindowLevel(window);
      },
      moveWindowTop: () => {
        window.moveTop();
      },
      onWindowsVisibleOverlayBlur:
        kind === 'visible' ? () => options.onVisibleWindowBlurred?.() : undefined,
    });
  });

  if (options.isDev && kind === 'visible') {
    window.webContents.openDevTools({ mode: 'detach' });
  }

  return window;
}

export function syncOverlayWindowLayer(window: BrowserWindow, layer: 'visible'): void {
  if (window.isDestroyed()) return;
  if (overlayWindowLayerByInstance.get(window) === layer) return;
  loadOverlayWindowLayer(window, layer);
}

export { buildOverlayWindowOptions } from './overlay-window-options';
export type { OverlayWindowKind } from './overlay-window-input';
