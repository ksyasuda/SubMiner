import os from 'node:os';
import path from 'node:path';
import type { BrowserWindow, BrowserWindowConstructorOptions } from 'electron';
import type { ConfigService } from '../../config/service';
import { promoteSettingsWindowAboveOverlay } from '../../core/services';
import type { Logger } from '../../logger';
import { getSyncHostsPath } from '../../shared/sync/sync-hosts-store';
import { createOpenConfigSettingsWindowHandler } from '../runtime/config-settings-window';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createCreateSyncUiWindowHandler } from '../runtime/setup-window-factory';
import { createSyncAutoScheduler } from '../runtime/sync-auto-scheduler';
import { resolveSyncLauncherCommand, runSyncLauncher } from '../runtime/sync-launcher-client';
import { createSyncUiRuntime, type SyncUiRuntimeDeps } from '../runtime/sync-ui-runtime';
import type { AppState } from '../state';

/**
 * What the sync window and auto-sync scheduler need from the rest of the app. Siblings built
 * after this assembly in `main.ts` are reached through getters or callbacks.
 */
export interface SyncAssemblyDeps {
  appState: Pick<AppState, 'syncUiWindow' | 'initialArgs'>;
  configService: Pick<ConfigService, 'getConfig'>;
  logger: Pick<Logger, 'info' | 'error'>;
  userDataPath: string;
  defaultImmersionDbPath: string;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showOverlayNotification'
  >;
  requestAppQuit: () => void;
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => BrowserWindow;
  /** Directory of the main bundle; the sync preload and page live next to it. */
  mainDirname: string;
  ipcMain: SyncUiRuntimeDeps['ipcMain'];
  /** Electron app details the launcher command depends on. */
  isPackaged: () => boolean;
  getAppPath: () => string;
  pickSnapshotFile: SyncUiRuntimeDeps['pickSnapshotFile'];
  revealPath: SyncUiRuntimeDeps['revealPath'];
}

export type SyncAssembly = ReturnType<typeof createSyncAssembly>;

/**
 * Owns the stats sync window, the sync IPC runtime, and the auto-sync scheduler. `main.ts`
 * registers the IPC handlers, starts the scheduler, and stops both on quit.
 */
export function createSyncAssembly(deps: SyncAssemblyDeps) {
  const { appState, configService, logger } = deps;

  const openSyncUiWindowHandler = createOpenConfigSettingsWindowHandler({
    getSettingsWindow: () => appState.syncUiWindow,
    setSettingsWindow: (window) => {
      appState.syncUiWindow = window as BrowserWindow | null;
    },
    createSettingsWindow: createCreateSyncUiWindowHandler({
      createBrowserWindow: (options) => deps.createBrowserWindow(options),
      preloadPath: path.join(deps.mainDirname, 'preload-syncui.js'),
    }),
    settingsHtmlPath: path.join(deps.mainDirname, 'syncui', 'index.html'),
    promoteSettingsWindowAboveOverlay: (window) =>
      promoteSettingsWindowAboveOverlay(window as BrowserWindow),
    onClosed: () => {
      // requestAppQuit (not app.quit) so the forced-exit fallback covers a
      // stalled quit; a standalone sync window is the app's only reason to live.
      if (appState.initialArgs?.syncWindow) deps.requestAppQuit();
    },
    log: (message) => logger.error(message),
  });

  const syncUiRuntime = createSyncUiRuntime({
    ipcMain: deps.ipcMain,
    hostsFilePath: getSyncHostsPath(deps.userDataPath),
    snapshotsDir: path.join(os.tmpdir(), 'subminer-db-snapshots'),
    getDbPath: () => {
      const configured = configService.getConfig().immersionTracking?.dbPath?.trim();
      return configured || deps.defaultImmersionDbPath;
    },
    resolveLauncherCommand: () =>
      resolveSyncLauncherCommand({
        execPath: process.execPath,
        appPath: deps.isPackaged() ? null : deps.getAppPath(),
      }),
    runLauncher: runSyncLauncher,
    getWindow: () => appState.syncUiWindow,
    pickSnapshotFile: deps.pickSnapshotFile,
    revealPath: deps.revealPath,
    nowMs: () => Date.now(),
    log: (message) => logger.info(message),
    notify: (payload) =>
      deps.getOverlayNotifications().showOverlayNotification({
        title: payload.title,
        body: payload.body,
        variant: payload.variant,
      }),
  });

  // Live sync uses a consistent WAL snapshot plus a transactional merge, and it
  // skips unfinished sessions, so resident stats writers do not block scheduling.
  const syncAutoScheduler = createSyncAutoScheduler({
    readState: () => syncUiRuntime.readState(),
    isRunning: () => syncUiRuntime.isRunning(),
    triggerHostSync: (host, direction) => {
      void syncUiRuntime.runHostSync({ host, direction }, { notify: true });
    },
    nowMs: () => Date.now(),
    log: (message) => logger.info(message),
  });

  return {
    openWindow: openSyncUiWindowHandler,
    runtime: syncUiRuntime,
    autoScheduler: syncAutoScheduler,
  };
}
