import { BrowserWindow, type Session } from 'electron';
import * as path from 'path';
import { WindowGeometry } from '../../types';
import { createLogger } from '../../logger';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  handleOverlayWindowBeforeInputEvent,
  type OverlayWindowKind,
} from './overlay-window-input';
import { buildOverlayWindowOptions } from './overlay-window-options';

const logger = createLogger('main:overlay-window');
const overlayWindowLayerByInstance = new WeakMap<BrowserWindow, OverlayWindowKind>();

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
  window.setBounds({
    x: geometry.x,
    y: geometry.y,
    width: geometry.width,
    height: geometry.height,
  });
}

export function ensureOverlayWindowLevel(window: BrowserWindow): void {
  if (process.platform === 'darwin') {
    window.setAlwaysOnTop(true, 'screen-saver', 1);
    window.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
    window.setFullScreenable(false);
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
    onWindowClosed: (kind: OverlayWindowKind) => void;
    yomitanSession?: Session | null;
  },
): BrowserWindow {
  const window = new BrowserWindow(buildOverlayWindowOptions(kind, options));

  options.ensureOverlayWindowLevel(window);
  loadOverlayWindowLayer(window, kind);

  window.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL) => {
    logger.error('Page failed to load:', errorCode, errorDescription, validatedURL);
  });

  window.webContents.on('did-finish-load', () => {
    options.onRuntimeOptionsChanged();
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
    if (!window.isDestroyed()) {
      options.ensureOverlayWindowLevel(window);
      if (kind === 'visible' && window.isVisible()) {
        window.moveTop();
      }
    }
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
