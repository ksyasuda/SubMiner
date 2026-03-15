import { BrowserWindow, ipcMain } from 'electron';
import * as path from 'path';
import type { WindowGeometry } from '../../types.js';
import { IPC_CHANNELS } from '../../shared/ipc/contracts.js';
import {
  buildStatsWindowLoadFileOptions,
  buildStatsWindowOptions,
  shouldHideStatsWindowForInput,
} from './stats-window-runtime.js';

let statsWindow: BrowserWindow | null = null;
let toggleRegistered = false;

export interface StatsWindowOptions {
  /** Absolute path to stats/dist/ directory */
  staticDir: string;
  /** Absolute path to the compiled preload-stats.js */
  preloadPath: string;
  /** Resolve the active stats toggle key from config */
  getToggleKey: () => string;
  /** Resolve the tracked overlay/mpv bounds */
  resolveBounds: () => WindowGeometry | null;
}

function syncStatsWindowBounds(window: BrowserWindow, bounds: WindowGeometry | null): void {
  if (!bounds || window.isDestroyed()) return;
  window.setBounds({
    x: bounds.x,
    y: bounds.y,
    width: bounds.width,
    height: bounds.height,
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

    const indexPath = path.join(options.staticDir, 'index.html');
    statsWindow.loadFile(indexPath, buildStatsWindowLoadFileOptions());

    statsWindow.on('closed', () => {
      statsWindow = null;
    });

    statsWindow.webContents.on('before-input-event', (event, input) => {
      if (shouldHideStatsWindowForInput(input, options.getToggleKey())) {
        event.preventDefault();
        statsWindow?.hide();
      }
    });

    statsWindow.once('ready-to-show', () => {
      if (statsWindow) {
        syncStatsWindowBounds(statsWindow, options.resolveBounds());
      }
      statsWindow?.show();
    });
  } else if (statsWindow.isVisible()) {
    statsWindow.hide();
  } else {
    syncStatsWindowBounds(statsWindow, options.resolveBounds());
    statsWindow.show();
    statsWindow.focus();
  }
}

/**
 * Register the IPC command handler for toggling the overlay.
 * Call this once during app initialization.
 */
export function registerStatsOverlayToggle(options: StatsWindowOptions): void {
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
