import { BrowserWindow } from 'electron';
import * as path from 'path';
import { WindowGeometry } from '../../types';
import { createLogger } from '../../logger';

const logger = createLogger('main:overlay-window');

export type OverlayWindowKind = 'visible' | 'invisible';

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
  window.setAlwaysOnTop(true);
}

export function enforceOverlayLayerOrder(options: {
  visibleOverlayVisible: boolean;
  invisibleOverlayVisible: boolean;
  mainWindow: BrowserWindow | null;
  invisibleWindow: BrowserWindow | null;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
}): void {
  if (!options.visibleOverlayVisible || !options.invisibleOverlayVisible) return;
  if (!options.mainWindow || options.mainWindow.isDestroyed()) return;
  if (!options.invisibleWindow || options.invisibleWindow.isDestroyed()) return;

  options.ensureOverlayWindowLevel(options.mainWindow);
  options.mainWindow.moveTop();
}

export function createOverlayWindow(
  kind: OverlayWindowKind,
  options: {
    isDev: boolean;
    overlayDebugVisualizationEnabled: boolean;
    ensureOverlayWindowLevel: (window: BrowserWindow) => void;
    onRuntimeOptionsChanged: () => void;
    setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
    isOverlayVisible: (kind: OverlayWindowKind) => boolean;
    tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
    onWindowClosed: (kind: OverlayWindowKind) => void;
  },
): BrowserWindow {
  const window = new BrowserWindow({
    show: false,
    width: 800,
    height: 600,
    x: 0,
    y: 0,
    transparent: true,
    frame: false,
    alwaysOnTop: true,
    skipTaskbar: true,
    resizable: false,
    hasShadow: false,
    focusable: true,
    webPreferences: {
      preload: path.join(__dirname, '..', '..', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      webSecurity: true,
      additionalArguments: [`--overlay-layer=${kind}`],
    },
  });

  options.ensureOverlayWindowLevel(window);

  const htmlPath = path.join(__dirname, '..', '..', 'renderer', 'index.html');

  window
    .loadFile(htmlPath, {
      query: { layer: kind === 'visible' ? 'visible' : 'invisible' },
    })
    .catch((err) => {
      logger.error('Failed to load HTML file:', err);
    });

  window.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL) => {
    logger.error('Page failed to load:', errorCode, errorDescription, validatedURL);
  });

  window.webContents.on('did-finish-load', () => {
    options.onRuntimeOptionsChanged();
    window.webContents.send(
      'overlay-debug-visualization:set',
      options.overlayDebugVisualizationEnabled,
    );
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
    if (!options.isOverlayVisible(kind)) return;
    if (!options.tryHandleOverlayShortcutLocalFallback(input)) return;
    event.preventDefault();
  });

  window.hide();

  window.on('closed', () => {
    options.onWindowClosed(kind);
  });

  window.on('blur', () => {
    if (!window.isDestroyed()) {
      options.ensureOverlayWindowLevel(window);
    }
  });

  if (options.isDev && kind === 'visible') {
    window.webContents.openDevTools({ mode: 'detach' });
  }

  return window;
}
