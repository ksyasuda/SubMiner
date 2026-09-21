import { BrowserWindow, dialog, ipcMain } from 'electron';
import { createLogger } from '../../logger.js';
import type { WindowGeometry } from '../../types.js';
import { IPC_CHANNELS } from '../../shared/ipc/contracts.js';
import {
  buildStatsWindowUrl,
  buildStatsWindowOptions,
  demoteVisibleStatsWindowBelowDialogs,
  presentStatsWindow,
  promoteStatsWindowLevel,
  promoteVisibleStatsWindowAboveOverlay,
  resolveStatsWindowOuterBoundsForContent,
  scheduleStatsWindowPostShowReconciles,
  showStatsNativeConfirmDialog,
  shouldHideStatsWindowForInput,
  shouldPresentStatsWindowAfterLoad,
  STATS_WINDOW_TITLE,
} from './stats-window-runtime.js';
import { ensureHyprlandWindowFloatingByTitle } from './hyprland-window-placement.js';
import {
  createStatsWindowLayerSuspensionState,
  isStatsWindowLayerSuspended,
  resetStatsWindowLayerSuspension,
  restoreStatsWindowLayer,
  suspendStatsWindowLayer,
} from './stats-window-layer.js';

let statsWindow: BrowserWindow | null = null;
let statsWindowGeneration = 0;
let toggleRegistered = false;
let nativeDialogLayerRegistered = false;
const nativeDialogLayerSuspension = createStatsWindowLayerSuspensionState();
const logger = createLogger('main:stats-window');

export interface StatsWindowOptions {
  /** Absolute path to the compiled preload-stats.js */
  preloadPath: string;
  /** Resolve the active stats API base URL */
  getApiBaseUrl: () => Promise<string> | string;
  /** Report server startup failure through the configured notification surface. */
  onStartupError?: (error: unknown) => void;
  /** Resolve the active stats toggle key from config */
  getToggleKey: () => string;
  /** Resolve the tracked overlay/mpv bounds */
  resolveBounds: () => WindowGeometry | null;
  /** Notify the main process when the stats overlay becomes visible/hidden */
  onVisibilityChanged?: (visible: boolean) => void;
}

function syncStatsWindowBounds(
  window: BrowserWindow,
  bounds: WindowGeometry | null,
): WindowGeometry | null {
  if (!bounds || window.isDestroyed()) return null;
  const outerBounds = resolveStatsWindowOuterBoundsForContent(window, bounds);
  window.setBounds({
    x: outerBounds.x,
    y: outerBounds.y,
    width: outerBounds.width,
    height: outerBounds.height,
  });
  return outerBounds;
}

function reconcileStatsWindowBounds(window: BrowserWindow, options: StatsWindowOptions): void {
  if (window.isDestroyed() || !window.isVisible()) {
    return;
  }
  const placementBounds = syncStatsWindowBounds(window, options.resolveBounds());
  if (placementBounds) {
    ensureHyprlandWindowFloatingByTitle({ title: STATS_WINDOW_TITLE, bounds: placementBounds });
  }
}

function scheduleStatsWindowBoundsReconcile(
  window: BrowserWindow,
  options: StatsWindowOptions,
): void {
  scheduleStatsWindowPostShowReconciles(() => {
    reconcileStatsWindowBounds(window, options);
  });
}

function showStatsWindow(window: BrowserWindow, options: StatsWindowOptions): void {
  const bounds = options.resolveBounds();
  let placementBounds = syncStatsWindowBounds(window, bounds);
  promoteStatsWindowLevel(window);
  presentStatsWindow(window);
  placementBounds = syncStatsWindowBounds(window, bounds) ?? placementBounds;
  if (
    !ensureHyprlandWindowFloatingByTitle({ title: STATS_WINDOW_TITLE, bounds: placementBounds })
  ) {
    placementBounds = syncStatsWindowBounds(window, bounds) ?? placementBounds;
  }
  options.onVisibilityChanged?.(true);
  promoteStatsOverlayAbovePlayback();
  reconcileStatsWindowBounds(window, options);
  scheduleStatsWindowBoundsReconcile(window, options);
}

export function promoteStatsOverlayAbovePlayback(): boolean {
  if (isStatsWindowLayerSuspended(nativeDialogLayerSuspension)) {
    return false;
  }

  if (!statsWindow) {
    return false;
  }

  return promoteVisibleStatsWindowAboveOverlay(statsWindow, {
    promoteHyprlandWindow: () => {
      ensureHyprlandWindowFloatingByTitle({ title: STATS_WINDOW_TITLE });
    },
  });
}

export function demoteStatsOverlayBelowDialogs(): boolean {
  if (!statsWindow) {
    return false;
  }

  return demoteVisibleStatsWindowBelowDialogs(statsWindow);
}

export function suspendStatsWindowLayerForNativeDialog(): void {
  if (!suspendStatsWindowLayer(nativeDialogLayerSuspension)) {
    return;
  }

  demoteStatsOverlayBelowDialogs();
}

export function restoreStatsWindowLayerAfterNativeDialog(): void {
  if (restoreStatsWindowLayer(nativeDialogLayerSuspension)) {
    promoteStatsOverlayAbovePlayback();
  }
}

function resetStatsWindowLayerAfterLifecycleEnd(): void {
  resetStatsWindowLayerSuspension(nativeDialogLayerSuspension);
}

export async function withStatsWindowLayerSuspendedForNativeDialog<T>(
  showDialog: () => Promise<T>,
): Promise<T> {
  suspendStatsWindowLayerForNativeDialog();
  try {
    return await showDialog();
  } finally {
    restoreStatsWindowLayerAfterNativeDialog();
  }
}

function confirmStatsNativeDialog(message: unknown): boolean {
  const dialogMessage =
    typeof message === 'string' && message.trim().length > 0 ? message : 'Confirm deletion?';

  return showStatsNativeConfirmDialog(statsWindow, dialogMessage, {
    showWithParent: (parentWindow, options) => dialog.showMessageBoxSync(parentWindow, options),
    showWithoutParent: (options) => dialog.showMessageBoxSync(options),
  });
}

function registerStatsNativeDialogLayerHandlers(): void {
  if (nativeDialogLayerRegistered) return;
  nativeDialogLayerRegistered = true;

  ipcMain.on(IPC_CHANNELS.command.statsNativeConfirmDialog, (event, message) => {
    event.returnValue = confirmStatsNativeDialog(message);
  });
  ipcMain.on(IPC_CHANNELS.command.statsNativeDialogOpened, (event) => {
    suspendStatsWindowLayerForNativeDialog();
    event.returnValue = true;
  });
  ipcMain.on(IPC_CHANNELS.command.statsNativeDialogClosed, () => {
    restoreStatsWindowLayerAfterNativeDialog();
  });
}

/**
 * Toggle the stats overlay window: create on first call, then show/hide.
 * The React app stays mounted across toggles — state is preserved.
 */
export async function toggleStatsOverlay(options: StatsWindowOptions): Promise<void> {
  if (!statsWindow) {
    const generation = statsWindowGeneration;
    const apiBaseUrl = await Promise.resolve()
      .then(() => options.getApiBaseUrl())
      .catch((error: unknown) => {
        options.onStartupError?.(error);
        throw error;
      });
    if (generation !== statsWindowGeneration || statsWindow) return;
    statsWindow = new BrowserWindow(
      buildStatsWindowOptions({
        preloadPath: options.preloadPath,
        bounds: options.resolveBounds(),
      }),
    );

    statsWindow.setTitle(STATS_WINDOW_TITLE);
    statsWindow.webContents.on('page-title-updated', (event) => {
      event.preventDefault();
      statsWindow?.setTitle(STATS_WINDOW_TITLE);
    });

    statsWindow.loadURL(buildStatsWindowUrl(apiBaseUrl));

    statsWindow.on('closed', () => {
      options.onVisibilityChanged?.(false);
      statsWindow = null;
      resetStatsWindowLayerAfterLifecycleEnd();
    });

    statsWindow.webContents.on('before-input-event', (event, input) => {
      if (shouldHideStatsWindowForInput(input, options.getToggleKey())) {
        event.preventDefault();
        statsWindow?.hide();
        options.onVisibilityChanged?.(false);
      }
    });
    const showInitialStatsWindow = () => {
      if (!statsWindow) return;
      showStatsWindow(statsWindow, options);
    };
    if (shouldPresentStatsWindowAfterLoad()) {
      statsWindow.webContents.once('did-finish-load', showInitialStatsWindow);
    } else {
      statsWindow.once('ready-to-show', showInitialStatsWindow);
    }

    statsWindow.on('blur', () => {
      if (!statsWindow || statsWindow.isDestroyed() || !statsWindow.isVisible()) {
        return;
      }
      promoteStatsOverlayAbovePlayback();
    });
  } else if (statsWindow.isVisible()) {
    statsWindow.hide();
    options.onVisibilityChanged?.(false);
  } else {
    showStatsWindow(statsWindow, options);
  }
}

/**
 * Register the IPC command handler for toggling the overlay.
 * Call this once during app initialization.
 */
export function registerStatsOverlayToggle(options: StatsWindowOptions): void {
  registerStatsNativeDialogLayerHandlers();
  if (toggleRegistered) return;
  toggleRegistered = true;
  ipcMain.on(IPC_CHANNELS.command.toggleStatsOverlay, () => {
    void toggleStatsOverlay(options).catch((error: unknown) => {
      logger.error('Failed to open stats overlay:', error);
    });
  });
}

/**
 * Clean up — destroy the stats window if it exists.
 * Call during app quit.
 */
export function destroyStatsWindow(): void {
  statsWindowGeneration += 1;
  if (statsWindow && !statsWindow.isDestroyed()) {
    statsWindow.destroy();
    statsWindow = null;
  }
  resetStatsWindowLayerAfterLifecycleEnd();
}
