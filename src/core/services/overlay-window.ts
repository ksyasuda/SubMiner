import electron from 'electron';
import type { BrowserWindow, Session } from 'electron';
import * as path from 'path';
import { WindowGeometry } from '../../types';
import { createLogger } from '../../logger';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  handleOverlayWindowBeforeInputEvent,
  handleOverlayWindowBlurred,
  type OverlayWindowKind,
} from './overlay-window-input';
import { ensureHyprlandWindowFloatingByTitle } from './hyprland-window-placement';
import { buildOverlayWindowOptions, OVERLAY_WINDOW_TITLES } from './overlay-window-options';
import { normalizeOverlayWindowBoundsForPlatform } from './overlay-window-bounds';
import { OVERLAY_WINDOW_CONTENT_READY_FLAG } from './overlay-window-flags';
export { OVERLAY_WINDOW_CONTENT_READY_FLAG } from './overlay-window-flags';

const logger = createLogger('main:overlay-window');
const { BrowserWindow: ElectronBrowserWindow, screen } = electron;
const overlayWindowLayerByInstance = new WeakMap<BrowserWindow, OverlayWindowKind>();
const overlayWindowContentReady = new WeakSet<BrowserWindow>();

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
  options: {
    promote?: boolean;
  } = {},
): void {
  if (!geometry || !window || window.isDestroyed()) return;
  const bounds = normalizeOverlayWindowBoundsForPlatform(
    geometry,
    process.platform,
    screen,
    window,
  );
  window.setBounds(bounds);
  ensureHyprlandWindowFloatingByTitle({
    title: window.getTitle(),
    bounds,
    promote: options.promote,
  });
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
  // Linux/X11 overlays start managed and only assert topmost while mpv owns the overlay layer.
  // Focus loss releases this again so native Wayland apps can cover the overlay on KDE.
  window.setAlwaysOnTop(true, 'screen-saver', 1);
  window.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
  ensureHyprlandWindowFloatingByTitle({ title: window.getTitle() });
  window.moveTop();
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
    linuxX11FullscreenOverlay?: boolean;
    onVisibleWindowBlurred?: () => void;
    onVisibleWindowFocused?: () => void;
    onWindowContentReady?: () => void;
    onWindowClosed: (kind: OverlayWindowKind, window: BrowserWindow) => void;
    yomitanSession?: Session | null;
  },
): BrowserWindow {
  const window = new ElectronBrowserWindow(buildOverlayWindowOptions(kind, options));
  window.setSkipTaskbar(true);
  (window as BrowserWindow & { [OVERLAY_WINDOW_CONTENT_READY_FLAG]?: boolean })[
    OVERLAY_WINDOW_CONTENT_READY_FLAG
  ] = false;

  if (!(process.platform === 'win32' && kind === 'visible')) {
    options.ensureOverlayWindowLevel(window);
  }
  loadOverlayWindowLayer(window, kind);

  window.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL) => {
    logger.error('Page failed to load:', errorCode, errorDescription, validatedURL);
  });

  window.webContents.on('did-finish-load', () => {
    window.setTitle(OVERLAY_WINDOW_TITLES[kind]);
    options.onRuntimeOptionsChanged();
  });

  window.webContents.on('page-title-updated', (event) => {
    event.preventDefault();
    window.setTitle(OVERLAY_WINDOW_TITLES[kind]);
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
    options.onWindowClosed(kind, window);
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
      onVisibleOverlayBlur:
        kind === 'visible' ? () => options.onVisibleWindowBlurred?.() : undefined,
    });
  });

  window.on('focus', () => {
    if (window.isDestroyed() || kind !== 'visible') return;
    options.onVisibleWindowFocused?.();
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
