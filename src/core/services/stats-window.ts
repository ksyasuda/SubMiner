import { BrowserWindow, dialog, ipcMain } from 'electron';
import * as path from 'path';
import type { WindowGeometry } from '../../types.js';
import { IPC_CHANNELS } from '../../shared/ipc/contracts.js';
import {
  buildStatsWindowLoadFileOptions,
  buildStatsWindowOptions,
  demoteVisibleStatsWindowBelowDialogs,
  presentStatsWindow,
  promoteStatsWindowLevel,
  promoteVisibleStatsWindowAboveOverlay,
  resolveStatsWindowOuterBoundsForContent,
  showStatsNativeConfirmDialog,
  shouldHideStatsWindowForInput,
  STATS_WINDOW_TITLE,
} from './stats-window-runtime.js';
import { ensureHyprlandWindowFloatingByTitle } from './hyprland-window-placement.js';

let statsWindow: BrowserWindow | null = null;
let toggleRegistered = false;
let nativeDialogLayerRegistered = false;
let nativeDialogLayerSuspensionCount = 0;

export interface StatsWindowOptions {
  /** Absolute path to stats/dist/ directory */
  staticDir: string;
  /** Absolute path to the compiled preload-stats.js */
  preloadPath: string;
  /** Resolve the active stats API base URL */
  getApiBaseUrl?: () => string;
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
}

export function promoteStatsOverlayAbovePlayback(): boolean {
  if (nativeDialogLayerSuspensionCount > 0) {
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
  nativeDialogLayerSuspensionCount += 1;
  if (nativeDialogLayerSuspensionCount !== 1) {
    return;
  }

  demoteStatsOverlayBelowDialogs();
}

export function restoreStatsWindowLayerAfterNativeDialog(): void {
  if (nativeDialogLayerSuspensionCount <= 0) {
    return;
  }

  nativeDialogLayerSuspensionCount -= 1;
  if (nativeDialogLayerSuspensionCount === 0) {
    promoteStatsOverlayAbovePlayback();
  }
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
export function toggleStatsOverlay(options: StatsWindowOptions): void {
  if (!statsWindow) {
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

    const indexPath = path.join(options.staticDir, 'index.html');
    statsWindow.loadFile(indexPath, buildStatsWindowLoadFileOptions(options.getApiBaseUrl?.()));

    statsWindow.on('closed', () => {
      options.onVisibilityChanged?.(false);
      statsWindow = null;
    });

    statsWindow.webContents.on('before-input-event', (event, input) => {
      if (shouldHideStatsWindowForInput(input, options.getToggleKey())) {
        event.preventDefault();
        statsWindow?.hide();
        options.onVisibilityChanged?.(false);
      }
    });
    statsWindow.once('ready-to-show', () => {
      if (!statsWindow) return;
      showStatsWindow(statsWindow, options);
    });

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
    toggleStatsOverlay(options);
  });
}

/**
 * Clean up — destroy the stats window if it exists.
 * Call during app quit.
 */
export function destroyStatsWindow(): void {
  if (statsWindow && !statsWindow.isDestroyed()) {
    statsWindow.destroy();
    statsWindow = null;
  }
}
