import fs from 'node:fs';
import path from 'node:path';
import type { BrowserWindow, BrowserWindowConstructorOptions } from 'electron';
import { AnkiConnectClient } from '../../anki-connect';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ConfigService } from '../../config/service';
import { buildConfigSettingsRegistry } from '../../config/settings/registry';
import {
  createConfigHotReloadRuntime,
  promoteSettingsWindowAboveOverlay,
} from '../../core/services';
import type { Logger } from '../../logger';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  buildConfigWarningDialogDetails,
  buildConfigWarningNotificationBody,
} from '../config-validation';
import {
  buildRestartRequiredConfigMessage,
  createConfigHotReloadAppliedHandler,
  createConfigHotReloadMessageHandler,
  createWatchConfigPathHandler,
} from '../runtime/config-hot-reload-handlers';
import {
  createConfigSettingsRuntime,
  type ConfigSettingsIpcMainLike,
} from '../runtime/config-settings-runtime';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createCreateConfigSettingsWindowHandler } from '../runtime/setup-window-factory';
import type { AppState } from '../state';

type ConfigHotReloadMessageDeps = Parameters<typeof createConfigHotReloadMessageHandler>[0];

/**
 * What config hot-reload and the settings window need from the rest of the app. Siblings built
 * after this assembly in `main.ts` are reached through getters or callbacks.
 */
export interface SettingsAssemblyDeps {
  appState: Pick<AppState, 'configSettingsWindow'>;
  configService: Pick<
    ConfigService,
    'getConfig' | 'getConfigPath' | 'getRawConfig' | 'getWarnings' | 'reloadConfigStrict'
  >;
  logger: Pick<Logger, 'error'>;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'getConfiguredStatusNotificationType' | 'showOverlayNotification'
  >;
  showMpvOsd: (message: string) => void;
  showDesktopNotification: ConfigHotReloadMessageDeps['showDesktopNotification'];
  /** Applies a reloaded config to the running app; shared by file hot-reload and the settings window. */
  hotReloadApplied: Parameters<typeof createConfigHotReloadAppliedHandler>[0];
  getYomitanAnkiDeckName: () => Promise<string>;
  showErrorBox: (title: string, content: string) => void;
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => BrowserWindow;
  /** Directory of the main bundle; the settings preload and page live next to it. */
  mainDirname: string;
  openPath: (targetPath: string) => Promise<string>;
  ipcMain: ConfigSettingsIpcMainLike;
}

export type SettingsAssembly = ReturnType<typeof createSettingsAssembly>;

/**
 * Owns config file hot-reload and the settings window, which both apply reloaded config through
 * the same handler. `main.ts` starts/stops hot-reload and registers the settings IPC handlers.
 */
export function createSettingsAssembly(deps: SettingsAssemblyDeps) {
  const { appState, configService, logger } = deps;
  const configSettingsFields = buildConfigSettingsRegistry(DEFAULT_CONFIG);

  const notifyConfigHotReloadMessage = createConfigHotReloadMessageHandler({
    getNotificationType: () => deps.getOverlayNotifications().getConfiguredStatusNotificationType(),
    showMpvOsd: (message) => deps.showMpvOsd(message),
    showOverlayNotification: (payload) =>
      deps.getOverlayNotifications().showOverlayNotification(payload),
    showDesktopNotification: (title, options) => deps.showDesktopNotification(title, options),
  });
  const watchConfigPathHandler = createWatchConfigPathHandler({
    fileExists: (targetPath) => fs.existsSync(targetPath),
    dirname: (targetPath) => path.dirname(targetPath),
    watchPath: (targetPath, listener) => fs.watch(targetPath, listener),
  });
  const applyConfigHotReloadDiff = createConfigHotReloadAppliedHandler(deps.hotReloadApplied);
  const configHotReloadRuntime = createConfigHotReloadRuntime({
    getCurrentConfig: () => configService.getConfig(),
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    watchConfigPath: (configPath, onChange) => watchConfigPathHandler(configPath, onChange),
    setTimeout: (callback, delayMs) => setTimeout(callback, delayMs),
    clearTimeout: (timeout) => clearTimeout(timeout),
    debounceMs: 250,
    onHotReloadApplied: applyConfigHotReloadDiff,
    onRestartRequired: (fields) =>
      notifyConfigHotReloadMessage(buildRestartRequiredConfigMessage(fields)),
    onInvalidConfig: notifyConfigHotReloadMessage,
    onValidationWarnings: (configPath, warnings) => {
      deps.showDesktopNotification('SubMiner', {
        body: buildConfigWarningNotificationBody(configPath, warnings),
      });
      if (process.platform === 'darwin') {
        deps.showErrorBox(
          'SubMiner config validation warning',
          buildConfigWarningDialogDetails(configPath, warnings),
        );
      }
    },
  });

  const configSettingsRuntime = createConfigSettingsRuntime({
    fields: configSettingsFields,
    getConfigPath: () => configService.getConfigPath(),
    getRawConfig: () => configService.getRawConfig(),
    getConfig: () => configService.getConfig(),
    getWarnings: () => configService.getWarnings(),
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    onHotReloadApplied: applyConfigHotReloadDiff,
    defaultAnkiConnectUrl: DEFAULT_CONFIG.ankiConnect.url,
    createAnkiClient: (url) => new AnkiConnectClient(url),
    getYomitanAnkiDeckName: deps.getYomitanAnkiDeckName,
    getSettingsWindow: () => appState.configSettingsWindow,
    setSettingsWindow: (window) => {
      appState.configSettingsWindow = window as BrowserWindow | null;
    },
    createSettingsWindow: createCreateConfigSettingsWindowHandler({
      createBrowserWindow: (options) => deps.createBrowserWindow(options),
      preloadPath: path.join(deps.mainDirname, 'preload-settings.js'),
    }),
    settingsHtmlPath: path.join(deps.mainDirname, 'settings', 'index.html'),
    promoteSettingsWindowAboveOverlay: (window) =>
      promoteSettingsWindowAboveOverlay(window as BrowserWindow),
    openPath: (targetPath) => deps.openPath(targetPath),
    ipcMain: deps.ipcMain,
    ipcChannels: IPC_CHANNELS.request,
    log: (message) => logger.error(message),
  });

  return {
    hotReload: configHotReloadRuntime,
    window: configSettingsRuntime,
  };
}
