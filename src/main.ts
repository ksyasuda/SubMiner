/*
  SubMiner - All-in-one sentence mining overlay
  Copyright (C) 2026 sudacode

  This program is free software: you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation, either version 3 of the License, or
  (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/
import {
  app,
  BrowserWindow,
  clipboard,
  globalShortcut,
  ipcMain,
  shell,
  protocol,
  Menu,
  nativeImage,
  Tray,
  dialog,
  screen,
} from 'electron';
import { applyControllerConfigUpdate } from './main/controller-config-update.js';
import { createHachidoriExtensionRuntime } from './core/services/hachidori-extension';
import { createDiscordRpcClient } from './main/runtime/discord-rpc-client.js';
import { startAppControlServer } from './main/runtime/app-control-server';
import { getAppControlSocketPath } from './shared/app-control';
import { restoreMacOSMpvFocusAfterModalClose } from './main/runtime/macos-modal-focus-handoff';
import {
  getDefaultPasswordStore,
  getPasswordStoreArg,
  normalizePasswordStoreArg,
} from './main/password-store-args';

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'chrome-extension',
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
      bypassCSP: true,
    },
  },
]);

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type {
  JimakuApiResponse,
  KikuFieldGroupingChoice,
  ResolvedConfig,
  SessionActionDispatchRequest,
  SubtitlePosition,
} from './types';
import { OPEN_ANKI_CARD_ACTION_ID } from './types';
import { downloadToFile, isRemoteMediaPath, parseMediaInfo } from './jimaku/utils';
import {
  createLogger,
  setLogFileToggles,
  setLogLevel,
  setLogRotation,
  type LogLevelSource,
} from './logger';
import {
  isLogFileEnabled,
  pruneLogDirectoryForPath,
  resolveDefaultLogFilePath,
} from './shared/log-files';
import { createFatalErrorReporter } from './main/fatal-error';
import {
  commandNeedsOverlayStartupPrereqs,
  commandNeedsOverlayRuntime,
  isHeadlessInitialCommand,
  parseArgs,
  shouldStartApp,
  type CliArgs,
  type CliCommandSource,
} from './cli/args';
import { printHelp } from './cli/help';
import { IPC_CHANNELS } from './shared/ipc/contracts';
import { buildMpvLoggingArgs } from './shared/mpv-logging-args';
import {
  getStartupModeFlags,
  shouldStartAutomaticUpdateChecks,
} from './main/runtime/startup-mode-flags';
import { buildConfigParseErrorDetails, failStartupFromConfig } from './main/config-validation';
import {
  findAnilistSetupDeepLinkArgvUrl,
  isAnilistTrackingEnabled,
} from './main/runtime/domains/anilist';
import { shouldShowTexthookerTrayEntry } from './main/runtime/tray-main-actions';
import { createGetDefaultSocketPathHandler } from './main/runtime/domains/jellyfin';
import {
  createGetRuntimeOptionsStateHandler,
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
  resolveSubtitleStyleForRenderer,
} from './main/runtime/domains/overlay';
import {
  createConfigDerivedRuntime,
  createMainSubsyncRuntime,
} from './main/runtime/domains/startup';
import { createMpvOsdRuntimeHandlers } from './main/runtime/domains/mpv';
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
  generateDefaultConfigFile,
  showDesktopNotification,
} from './core/utils';
import {
  SubtitleWebSocket,
  Texthooker,
  createDiscordPresenceService,
  createOverlayContentMeasurementStore,
  createOverlayManager,
  createOverlayWindow as createOverlayWindowCore,
  hasMpvWebsocketPlugin,
  loadSubtitlePosition as loadSubtitlePositionCore,
  promoteSettingsWindowAboveOverlay,
  runStartupBootstrapRuntime,
  saveSubtitlePosition as saveSubtitlePositionCore,
  showMpvOsdRuntime,
} from './core/services';
import {
  destroyStatsWindow,
  promoteStatsOverlayAbovePlayback,
  registerStatsOverlayToggle,
  toggleStatsOverlay as toggleStatsOverlayWindow,
  withStatsWindowLayerSuspendedForNativeDialog,
} from './core/services/stats-window.js';
import { createCharacterDictionaryAssembly } from './main/assemblies/character-dictionary';
import { createDictionariesAssembly } from './main/assemblies/dictionaries';
import { createJellyfinAssembly } from './main/assemblies/jellyfin';
import { createAnilistAssembly } from './main/assemblies/anilist';
import { createSettingsAssembly } from './main/assemblies/settings';
import { createSyncAssembly } from './main/assemblies/sync';
import { createYoutubeAssembly } from './main/assemblies/youtube';
import { createOverlayModalsAssembly } from './main/assemblies/overlay-modals';
import { createFirstRunSetupAssembly } from './main/assemblies/first-run-setup';
import { createOverlayAssembly } from './main/assemblies/overlay';
import { createYomitanAssembly } from './main/assemblies/yomitan';
import { createStatsAssembly } from './main/assemblies/stats';
import { createMiningAssembly } from './main/assemblies/mining';
import { createShortcutsAssembly } from './main/assemblies/shortcuts';
import { createTrayAssembly } from './main/assemblies/tray';
import { createPluginsAssembly } from './main/assemblies/plugins';
import { createSessionActionsAssembly } from './main/assemblies/session-actions';
import { createSubtitleGenerationAssembly } from './main/assemblies/subtitle-generation';
import { createStartupAssembly } from './main/assemblies/startup';
import type { IpcHandlerDeps } from './main/assemblies/slices';
import type { CliCommandContextMainDeps } from './main/runtime/cli-command-context-factory';
import { createPlaybackAssembly } from './main/assemblies/playback';
import { createWindowsMpvPluginDetectionRuntime } from './main/runtime/windows-mpv-plugin-detection';
import { refreshManagedCommandLineLauncher } from './main/runtime/command-line-launcher';
import { applyLoginShellPath } from './main/runtime/login-shell-path';
import { createWindowsMpvLaunchDeps, launchWindowsMpv } from './main/runtime/windows-mpv-launch';
import {
  isJellyfinConfiguredForTray as isJellyfinConfiguredForTrayRuntime,
  toggleJellyfinDiscoveryFromTray as toggleJellyfinDiscoveryFromTrayRuntime,
} from './main/runtime/jellyfin-tray-discovery';
import {
  shouldEnsureTrayOnStartupForInitialArgs,
  shouldQuitOnWindowAllClosedForTrayState,
} from './main/runtime/startup-tray-policy';
import { createLogExportTrayRuntime } from './main/runtime/log-export-tray';
import { createForceQuitHandler } from './main/runtime/app-lifecycle-actions';
import { createAnilistUpdateQueue } from './core/services/anilist/anilist-update-queue';
import { createAnilistRateLimiter } from './core/services/anilist/rate-limiter';
import { createJellyfinTokenStore } from './core/services/jellyfin-token-store';
import { createAnilistTokenStore } from './core/services/anilist/anilist-token-store';
import {
  composeCliStartupHandlers,
  composeHeadlessStartupHandlers,
  composeIpcRuntimeHandlers,
  composeStartupLifecycleHandlers,
} from './main/runtime/composers';
import { createStartupBootstrapRuntimeDeps } from './main/startup';
import { createAppLifecycleRuntimeRunner } from './main/startup-lifecycle';
import {
  registerSecondInstanceHandlerEarly,
  requestSingleInstanceLockEarly,
  shouldBypassSingleInstanceLockForArgv,
} from './main/early-single-instance';
import { handleMpvCommandFromIpcRuntime } from './main/ipc-mpv-command';
import { registerIpcRuntimeServices } from './main/ipc-runtime';
import {
  createSubtitleSelectionRuntime,
  registerSubtitleSelectionIpc,
} from './main/runtime/subtitle-selection';
import { createAnkiJimakuIpcRuntimeServiceDeps } from './main/dependencies';
import { createMainBootServices, type MainBootServicesResult } from './main/boot/services';
import { handleCliCommandRuntimeServiceWithContext } from './main/cli-runtime';
import { createOverlayModalRuntimeService } from './main/overlay-runtime';
import { createOverlayModalInputState } from './main/runtime/overlay-modal-input-state';
import { getSharedRemoteMediaWindowCache } from './core/services/remote-media-window-cache';
import { openSubsyncManualModal as openSubsyncManualModalRuntime } from './main/runtime/subsync-open';
import { createChangelogRuntime } from './main/runtime/changelog/changelog-runtime';
import { PLUGIN_DEACTIVATE_TIMEOUT_MS } from './core/services/plugins/plugin-host';
import { createPlaylistBrowserIpcRuntime } from './main/runtime/playlist-browser-ipc';
import { createMediaRuntimeService } from './main/media-runtime';
import { createDiscordPresenceRuntime } from './main/runtime/discord-presence-runtime';
import { createStartupOsdSequencer } from './main/runtime/startup-osd-sequencer';
import {
  INSTALL_UPDATE_ACTION_ID,
  UPDATE_AVAILABLE_NOTIFICATION_ID,
  VIEW_CHANGELOG_ACTION_ID,
} from './main/runtime/update/update-notifications';
import { createOverlayNotificationsRuntime } from './main/runtime/overlay-notifications-runtime';
import {
  runUpdateCliCommand,
  writeUpdateCliCommandResponse,
} from './main/runtime/update/update-cli-command';
import {
  runEnsureLinuxRuntimePluginAssetsCliCommand,
  writeEnsureLinuxRuntimePluginAssetsCliCommandResponse,
} from './main/runtime/linux-runtime-plugin-assets-cli-command';
import { ensureLinuxRuntimePluginAssets } from './main/runtime/linux-runtime-plugin-assets';
import { createUpdateServiceRuntime } from './main/runtime/update/update-service-runtime';
import { shouldSuppressVisibleOverlayRaiseForSeparateWindow } from './main/runtime/settings-window-z-order';
import { type StartupState, applyStartupState, createAppState } from './main/state';
import { ConfigService, DEFAULT_CONFIG, generateConfigTemplate } from './config';
import { resolveConfigDir } from './config/path-resolution';
import { createSubtitlePrefetchService } from './core/services/subtitle-prefetch';
import { loadSubtitleSourceText } from './main/runtime/internal-subtitle-extraction';

if (process.platform === 'linux') {
  app.commandLine.appendSwitch('enable-features', 'GlobalShortcutsPortal');
  const passwordStore = normalizePasswordStoreArg(
    getPasswordStoreArg(process.argv) ?? getDefaultPasswordStore(),
  );
  app.commandLine.appendSwitch('password-store', passwordStore);
  createLogger('main').debug(`Applied --password-store ${passwordStore}`);
}

app.setName('SubMiner');

const DEFAULT_TEXTHOOKER_PORT = 5174;
const DEFAULT_MPV_LOG_FILE = resolveDefaultLogFilePath('mpv', {
  platform: process.platform,
  homeDir: os.homedir(),
  appDataDir: process.env.APPDATA,
});

const DISCORD_PRESENCE_APP_ID = '1475264834730856619';

let notifyAnilistTokenStoreWarning: (message: string) => void = () => {};

const isDev = process.argv.includes('--dev') || process.argv.includes('--debug');
const texthookerService = new Texthooker(() => {
  const config = configService.getConfig();
  const characterDictionaryEnabled =
    config.subtitleStyle.nameMatchEnabled && yomitan.profilePolicy.isCharacterDictionaryEnabled();
  const knownWordColoringEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.knownWords.highlightEnabled',
    config.ankiConnect.knownWords.highlightEnabled,
  );
  const nPlusOneColoringEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.nPlusOne',
    config.ankiConnect.nPlusOne.enabled,
  );

  return {
    enableKnownWordColoring: knownWordColoringEnabled,
    enableNPlusOneColoring: nPlusOneColoringEnabled,
    enableNameMatchColoring: config.subtitleStyle.nameMatchEnabled && characterDictionaryEnabled,
    enableFrequencyColoring: getRuntimeBooleanOption(
      'subtitle.annotation.frequency',
      config.subtitleStyle.frequencyDictionary.enabled,
    ),
    enableJlptColoring: getRuntimeBooleanOption(
      'subtitle.annotation.jlpt',
      config.subtitleStyle.enableJlpt,
    ),
    characterDictionaryEnabled,
    knownWordColor: config.subtitleStyle.knownWordColor,
    nPlusOneColor: config.subtitleStyle.nPlusOneColor,
    nameMatchColor: config.subtitleStyle.nameMatchColor,
    hoverTokenColor: config.subtitleStyle.hoverTokenColor,
    hoverTokenBackgroundColor: config.subtitleStyle.hoverTokenBackgroundColor,
    jlptColors: config.subtitleStyle.jlptColors,
    frequencyDictionary: {
      singleColor: config.subtitleStyle.frequencyDictionary.singleColor,
      bandedColors: config.subtitleStyle.frequencyDictionary.bandedColors,
    },
  };
});
const getDefaultSocketPathHandler = createGetDefaultSocketPathHandler({
  platform: process.platform,
});

type BootServices = MainBootServicesResult<
  ConfigService,
  ReturnType<typeof createAnilistTokenStore>,
  ReturnType<typeof createJellyfinTokenStore>,
  ReturnType<typeof createAnilistUpdateQueue>,
  SubtitleWebSocket,
  ReturnType<typeof createLogger>,
  ReturnType<typeof createOverlayManager>,
  ReturnType<typeof createOverlayModalInputState>,
  ReturnType<typeof createOverlayContentMeasurementStore>,
  ReturnType<typeof createOverlayModalRuntimeService>,
  ReturnType<typeof createAppState>,
  {
    requestSingleInstanceLock: () => boolean;
    quit: () => void;
    exit: (code?: number) => void;
    on: (event: string, listener: (...args: unknown[]) => void) => unknown;
    whenReady: () => Promise<void>;
  }
>;

const bootServices = createMainBootServices({
  platform: process.platform,
  argv: process.argv,
  configDir: app.getPath('userData'),
  appDataDir: process.env.APPDATA,
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  homeDir: os.homedir(),
  defaultMpvLogFile: DEFAULT_MPV_LOG_FILE,
  envMpvLog: process.env.SUBMINER_MPV_LOG,
  defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
  getDefaultSocketPath: () => getDefaultSocketPathHandler(),
  resolveConfigDir,
  existsSync: fs.existsSync,
  mkdirSync: fs.mkdirSync,
  joinPath: (...parts) => path.join(...parts),
  app,
  shouldBypassSingleInstanceLock: () => shouldBypassSingleInstanceLockForArgv(process.argv),
  requestSingleInstanceLockEarly: () => requestSingleInstanceLockEarly(app),
  registerSecondInstanceHandlerEarly: (listener) => {
    registerSecondInstanceHandlerEarly(app, listener);
  },
  onConfigStartupParseError: (error) => {
    failStartupFromConfig(
      'SubMiner config parse error',
      buildConfigParseErrorDetails(error.path, error.parseError),
      {
        logError: (details) => console.error(details),
        showErrorBox: (title, details) => dialog.showErrorBox(title, details),
        quit: () => requestAppQuit(),
      },
    );
  },
  createConfigService: (configDir) => new ConfigService(configDir),
  createAnilistTokenStore: (targetPath) =>
    createAnilistTokenStore(targetPath, {
      info: (message: string) => console.info(message),
      warn: (message: string, details?: unknown) => console.warn(message, details),
      error: (message: string, details?: unknown) => console.error(message, details),
      warnUser: (message: string) => notifyAnilistTokenStoreWarning(message),
    }),
  createJellyfinTokenStore: (targetPath) =>
    createJellyfinTokenStore(targetPath, {
      info: (message: string) => console.info(message),
      warn: (message: string, details?: unknown) => console.warn(message, details),
      error: (message: string, details?: unknown) => console.error(message, details),
    }),
  createAnilistUpdateQueue: (targetPath) =>
    createAnilistUpdateQueue(targetPath, {
      info: (message: string) => console.info(message),
      warn: (message: string, details?: unknown) => console.warn(message, details),
      error: (message: string, details?: unknown) => console.error(message, details),
    }),
  createSubtitleWebSocket: (payloadMode) => new SubtitleWebSocket(payloadMode),
  createLogger,
  createOverlayManager: () =>
    createOverlayManager({
      shouldPromoteWindowOnBoundsUpdate: (window) =>
        !shouldSuppressVisibleOverlayRaiseForSeparateWindow({
          window,
          mainWindow: overlayManager.getMainWindow(),
          separateWindows: getOverlayForegroundSeparateWindows(),
        }),
    }),
  createOverlayModalInputState,
  createOverlayContentMeasurementStore: ({ logger }) => {
    return createOverlayContentMeasurementStore({
      now: () => Date.now(),
      warn: (message: string) => logger.warn(message),
    });
  },
  // Read lazily from modal-state callbacks, which only fire once the assemblies below exist.
  getSyncOverlayShortcutsForModal: () => (isActive) =>
    shortcuts.syncOverlayShortcutsForModal(isActive),
  getSyncOverlayVisibilityForModal: () => () =>
    overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
  createOverlayModalRuntime: ({ overlayManager, overlayModalInputState }) => {
    return createOverlayModalRuntimeService(
      {
        getMainWindow: () => overlayManager.getMainWindow(),
        getModalWindow: () => overlayManager.getModalWindow(),
        createModalWindow: () => createModalWindow(),
        getModalGeometry: () => overlay.geometry.getCurrentOverlayGeometry(),
        setModalWindowBounds: (geometry) => overlayManager.setModalWindowBounds(geometry),
      },
      {
        onModalStateChange: (isActive: boolean) =>
          overlayModalInputState.handleModalInputStateChange(isActive),
        // On macOS, after the last modal closes the post-close visibility sync hides the overlay
        // when neither it nor mpv is focused, and keyboard focus is left in limbo (mpv keys like
        // shift-to-unpause stop working). Programmatically focusing mpv from our background helper
        // is refused by macOS (most visibly in native fullscreen), so instead resign SubMiner's
        // active status — exactly what a manual click does — handing focus back to the previously
        // active app (mpv). The overlay is already hidden at this point, so app.hide() hides nothing
        // visible; once mpv is focused the tracker re-shows the overlay above it.
        onFinalModalClosed: () => {
          void restoreMacOSMpvFocusAfterModalClose({
            platform: process.platform,
            focusMpv: async () => {
              app.hide();
            },
            getWindowTracker: () => appState.windowTracker,
            updateVisibleOverlayVisibility: () =>
              overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
            warn: (message, details) => logger.warn(message, details),
          });
        },
      },
    );
  },
  createAppState,
}) as BootServices;
const {
  configDir: CONFIG_DIR,
  userDataPath: USER_DATA_PATH,
  defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  configService,
  anilistTokenStore,
  jellyfinTokenStore,
  anilistUpdateQueue,
  subtitleWsService,
  annotationSubtitleWsService,
  logger,
  overlayManager,
  overlayModalInputState,
  overlayContentMeasurementStore,
  overlayModalRuntime,
  appState,
  appLifecycleApp,
} = bootServices;
// Backend changes take effect on restart; windows and parser must share one session.
const activeDictionaryBackend = configService.getConfig().dictionaryBackend;
const yomitan = createYomitanAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  userDataPath: USER_DATA_PATH,
  dictionaryBackend: activeDictionaryBackend,
  hachidori: createHachidoriExtensionRuntime(USER_DATA_PATH),
  getOverlayWindows: () => overlayManager.getOverlayWindows(),
  getOverlayNotifications: () => overlayNotificationsRuntime,
  ipcMain,
  openExternal: (url) => shell.openExternal(url),
});
yomitan.registerDictionaryExternalLinkHandler();

function getOverlayForegroundSeparateWindows(): BrowserWindow[] {
  return [
    appState.configSettingsWindow,
    appState.syncUiWindow,
    ...yomitan.getSettingsWindows(),
    appState.anilistSetupWindow,
    appState.jellyfinSetupWindow,
    appState.firstRunSetupWindow,
  ].filter((window): window is BrowserWindow => Boolean(window));
}

notifyAnilistTokenStoreWarning = (message: string) => {
  logger.warn(`[AniList] ${message}`);
  try {
    showDesktopNotification('SubMiner AniList', {
      body: message,
    });
  } catch {
    // Notification may fail if desktop notifications are unavailable early in startup.
  }
};
const appLogger = {
  logInfo: (message: string) => {
    logger.info(message);
  },
  logDebug: (message: string) => {
    logger.debug(message);
  },
  logWarning: (message: string) => {
    logger.warn(message);
  },
  logError: (message: string, details: unknown) => {
    logger.error(message, details);
  },
  logNoRunningInstance: () => {
    logger.error('No running instance. Use --start to launch the app.');
  },
  logConfigWarning: (warning: {
    path: string;
    message: string;
    value: unknown;
    fallback: unknown;
  }) => {
    logger.warn(
      `[config] ${warning.path}: ${warning.message} value=${JSON.stringify(warning.value)} fallback=${JSON.stringify(warning.fallback)}`,
    );
  },
};
const reportFatalError = createFatalErrorReporter({
  showErrorBox: (title, details) => dialog.showErrorBox(title, details),
  consoleError: (message, error) => logger.error(message, error),
});

let forceQuitTimer: ReturnType<typeof setTimeout> | null = null;
const anilistRateLimiter = createAnilistRateLimiter();
const stats = createStatsAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  userDataPath: USER_DATA_PATH,
  mainDirname: __dirname,
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  anilistRateLimiter,
  statsWindow: {
    registerToggle: (options) => registerStatsOverlayToggle(options),
    toggle: (options) => toggleStatsOverlayWindow(options),
  },
  yomitanParserRuntimeDeps: yomitan.getParserRuntimeDeps(),
  getYomitanAnkiDeckName: () => yomitan.getCurrentAnkiDeckName(),
  getTokenizeSubtitle: () => getTokenizeSubtitleDeferred(),
  createMecabTokenizerAndCheck: () => createMecabTokenizerAndCheck(),
  getOverlayNotifications: () => overlayNotificationsRuntime,
  getOverlayBounds: () => overlay.geometry.getCurrentOverlayGeometry(),
  onStatsOverlayVisibilityChanged: (visible) =>
    overlay.interaction.handleStatsOverlayVisibilityChanged(visible),
  runAnilistPostWatchUpdate: () => anilist.maybeRunAnilistPostWatchUpdate({ force: true }),
  openExternal: (url) => shell.openExternal(url),
  requestAppQuit: () => requestAppQuit(),
});

function requestAppQuit(): void {
  yomitan.destroySettingsWindows();
  destroyStatsWindow();
  void stats.stopStatsServer().catch((error: unknown) => {
    logger.warn('Failed to stop stats server while quitting.', error);
  });
  if (!forceQuitTimer) {
    // Quit cleanup unloads plugins first, so allow their budget plus 2 s for the rest.
    forceQuitTimer = setTimeout(() => {
      logger.warn('App quit timed out; forcing process exit.');
      void createForceQuitHandler({
        destroyImmersionTracker: () => appState.immersionTracker?.destroy(),
        logError: (error) => logger.error('Failed to finalize stats before forced exit.', error),
        exit: () => app.exit(0),
      })();
    }, PLUGIN_DEACTIVATE_TIMEOUT_MS + 2000);
  }
  app.quit();
}

process.on('SIGINT', () => {
  requestAppQuit();
});
process.on('SIGTERM', () => {
  requestAppQuit();
});

// Managed Whisper models live next to the config file.
const getSubtitleGenerationModelDirectory = (): string =>
  path.join(path.dirname(configService.getConfigPath()), 'models', 'whisper');
const youtube = createYoutubeAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  getWhisperModelDirectory: getSubtitleGenerationModelDirectory,
  getMainOverlayWindow: () => overlayManager.getMainWindow(),
  overlayModalRuntime,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  getOverlayGeometry: () => overlay.geometry,
  showMpvOsd: (text) => showMpvOsd(text),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  refreshCurrentSubtitle: (text: string) => {
    subtitleProcessingController.refreshCurrentSubtitle(text);
  },
  refreshSubtitleSidebarSource: async (sourcePath: string, mediaPath?: string) => {
    await subtitlePrefetchRuntime.refreshSubtitleSidebarFromSource(sourcePath, mediaPath);
  },
  startTokenizationWarmups: async () => {
    await startTokenizationWarmups();
  },
  waitForMediaTokenizationReady: (mediaPath) =>
    playback.currentMediaTokenizationGate.waitUntilReady(mediaPath),
  generateYoutubeSubtitles: (request) =>
    subtitleGeneration.runtime.generateYoutubeSubtitles(request),
  openSubtitleGenerationModal: () => subtitleGeneration.openModal(),
  invalidatePendingAutoplayReadyFallbacks: () =>
    playback.autoplayReadyGate.invalidatePendingAutoplayReadyFallbacks(),
  ensurePlaybackRuntimeReady: () => startup.ensureYoutubePlaybackRuntimeReady(),
  launchWindowsMpv: (playbackUrl, args) => {
    const config = configService.getConfig();
    setLogFileToggles(config.logging.files);
    const mpvLogPath = isLogFileEnabled('mpv') ? resolveDefaultLogFilePath('mpv') : '';
    if (mpvLogPath) {
      pruneLogDirectoryForPath(mpvLogPath, config.logging.rotation);
    }
    const mpvArgs = [...args, ...buildMpvLoggingArgs(config.logging.level, mpvLogPath, args)];
    return launchWindowsMpv(
      [playbackUrl],
      createWindowsMpvLaunchDeps({
        showError: (title, content) => dialog.showErrorBox(title, content),
      }),
      mpvArgs,
      process.execPath,
      resolveBundledMpvRuntimePluginEntrypoint(),
      config.mpv.executablePath,
      config.mpv.launchMode,
      {
        detectInstalledMpvPlugin: detectWindowsInstalledMpvPlugin,
        notifyInstalledPluginDetected: logInstalledMpvPluginDetected,
        resolveInstalledPluginBeforeLaunch: (detection, mpvPath) =>
          promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(mpvPath, detection),
      },
      getMpvPluginRuntimeConfig(),
    );
  },
  ensureMpvConnectedForPlayback: () => jellyfin.ensureMpvConnectedForPlayback(),
  requestAppQuit: () => requestAppQuit(),
  browserWindow: {
    createBrowserWindow: (options) => new BrowserWindow(options),
    preloadPath: path.join(__dirname, 'preload-youtube-browser.js'),
    ipcMain,
    showContextMenu: (window, template) => {
      Menu.buildFromTemplate(template).popup({ window });
    },
    openExternal: (url) => {
      void shell.openExternal(url);
    },
  },
});
const {
  resolveBundledMpvRuntimePluginEntrypoint,
  detectWindowsInstalledMpvPlugin,
  logInstalledMpvPluginDetected,
  promptForLegacyMpvPluginRemovalBeforeWindowsLaunch,
} = createWindowsMpvPluginDetectionRuntime({
  mainDirname: __dirname,
  logWarn: (message) => logger.warn(message),
});

function getMpvPluginRuntimeConfig() {
  const config = configService.getConfig();
  return {
    socketPath: appState.mpvSocketPath,
    binaryPath: config.mpv.subminerBinaryPath,
    backend: config.mpv.backend,
    autoStart: config.mpv.autoStartSubMiner,
    autoStartVisibleOverlay: config.auto_start_overlay,
    autoStartPauseUntilReady: config.mpv.pauseUntilOverlayReady,
    osdMessages:
      config.ankiConnect.behavior.notificationType === 'osd' ||
      config.ankiConnect.behavior.notificationType === 'osd-system',
    texthookerEnabled: config.texthooker.launchAtStartup,
  };
}

// Finder/Dock launches inherit launchd's minimal PATH; pick up the user's shell PATH so
// launcher/Bun detection and spawned tools match what their terminal sees.
const loginShellPathReady: Promise<void> =
  process.platform === 'darwin'
    ? applyLoginShellPath({ env: process.env })
        .then((applied) => {
          if (!applied) logger.warn('Login shell PATH was empty; using inherited PATH');
        })
        .catch((error) => {
          logger.warn('Failed to read login shell PATH; using inherited PATH', error);
        })
    : Promise.resolve();
const createCommandLineLauncherRuntimeOptions = async () => {
  await loginShellPathReady;
  return {
    platform: process.platform,
    env: process.env,
    homeDir: os.homedir(),
    localAppData: process.env.LOCALAPPDATA,
    userProfile: process.env.USERPROFILE,
    cwd: process.cwd(),
    resourcesPath: process.resourcesPath,
    appExePath: process.execPath,
    appVersion: app.getVersion(),
    bundledBunPath: app.isPackaged
      ? path.join(process.resourcesPath, 'bun', process.platform === 'win32' ? 'bun.exe' : 'bun')
      : undefined,
  };
};
const firstRunSetup = createFirstRunSetupAssembly({
  appState,
  configService,
  logger,
  configDir: CONFIG_DIR,
  userDataPath: USER_DATA_PATH,
  dictionaryBackend: activeDictionaryBackend,
  ensureYomitanExtensionLoaded: () => yomitan.ensureExtensionLoaded(),
  getYomitanParserRuntimeDeps: () => yomitan.getParserRuntimeDeps(),
  createCommandLineLauncherRuntimeOptions: () => createCommandLineLauncherRuntimeOptions(),
  getAppPath: (name) => app.getPath(name),
  trashItem: (targetPath) => shell.trashItem(targetPath),
  writeShortcutLink: (shortcutPath, operation, details) =>
    shell.writeShortcutLink(shortcutPath, operation, details),
  createBrowserWindow: (options) => new BrowserWindow(options),
  openDictionarySettings: () => yomitan.openDictionarySettings(),
  openConfigSettingsWindow: () => settings.window.openWindow(),
  onSetupStateChanged: () => tray.refreshMenu(),
  requestAppQuit: () => requestAppQuit(),
});
const discordPresenceSessionStartedAtMs = Date.now();
let discordPresenceMediaDurationSec: number | null = null;
const discordPresenceRuntime = createDiscordPresenceRuntime({
  getDiscordPresenceService: () => appState.discordPresenceService,
  isDiscordPresenceEnabled: () => configService.getConfig().discordPresence.enabled === true,
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentSubtitleText: () => appState.currentSubText,
  getPlaybackPaused: () => appState.playbackPaused,
  getFallbackMediaDurationSec: () => anilist.getMediaDurationSec(),
  getSessionStartedAtMs: () => discordPresenceSessionStartedAtMs,
  getMediaDurationSec: () => discordPresenceMediaDurationSec,
  setMediaDurationSec: (next) => {
    discordPresenceMediaDurationSec = next;
  },
});

async function initializeDiscordPresenceService(): Promise<void> {
  if (configService.getConfig().discordPresence.enabled !== true) {
    appState.discordPresenceService = null;
    return;
  }

  appState.discordPresenceService = createDiscordPresenceService({
    config: configService.getConfig().discordPresence,
    createClient: () => createDiscordRpcClient(DISCORD_PRESENCE_APP_ID),
    logDebug: (message, meta) => logger.debug(message, meta),
  });
  await appState.discordPresenceService.start();
  discordPresenceRuntime.publishDiscordPresence();
}
const configDerivedRuntime = createConfigDerivedRuntime({
  getResolvedConfig: () => configService.getConfig(),
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
  defaultJimakuLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  defaultJimakuMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
  defaultJimakuApiBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
});
const subsyncRuntime = createMainSubsyncRuntime({
  getMpvClient: () => appState.mpvClient,
  getResolvedConfig: () => configService.getConfig(),
  getSubsyncInProgress: () => appState.subsyncInProgress,
  setSubsyncInProgress: (inProgress) => {
    appState.subsyncInProgress = inProgress;
  },
  showMpvOsd: (text) => overlayNotificationsRuntime.showSubsyncStatusNotification(text),
  openManualPicker: (payload) => {
    overlayModals.openOverlayHostedModalWithOsd(
      (deps) => openSubsyncManualModalRuntime(deps, payload),
      'Subsync overlay unavailable.',
      'Failed to open subsync overlay.',
    );
  },
});
const startupOsdSequencer = createStartupOsdSequencer({
  getNotificationType: () => overlayNotificationsRuntime.getConfiguredStatusNotificationType(),
  showOsd: (message) => showMpvOsd(message),
  showOverlayNotification: (payload) =>
    overlayNotificationsRuntime.showOverlayNotification(payload),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
});
const tray = createTrayAssembly({
  appState,
  logger,
  iconRoots: {
    resourcesPath: process.resourcesPath,
    appPath: app.getAppPath(),
    mainDirname: __dirname,
  },
  electron: {
    createTray: (icon) => new Tray(icon),
    createImageFromPath: (iconPath) => nativeImage.createFromPath(iconPath),
    createEmptyImage: () => nativeImage.createEmpty(),
    buildMenuFromTemplate: (template) => Menu.buildFromTemplate(template),
  },
  menu: {
    openSessionHelpModal: () => overlayModals.openSessionHelpOverlay(),
    openChangelogModal: () => overlayModals.openChangelogOverlay(),
    openTexthookerInBrowser: () => handleCliCommand(parseArgs(['--texthooker', '--open-browser'])),
    showTexthookerPage: () => shouldShowTexthookerTrayEntry(configService.getConfig()),
    showFirstRunSetup: () => !firstRunSetup.service.isSetupCompleted(),
    openFirstRunSetupWindow: (force?: boolean) => firstRunSetup.openWindow(force),
    getDictionaryBackend: () => activeDictionaryBackend,
    openHachidoriSettings: () => yomitan.openHachidoriSettings(),
    openYomitanSettings: () => yomitan.openYomitanSettings(),
    openConfigSettingsWindow: () => settings.window.openWindow(),
    openSyncUiWindow: () => sync.openWindow(),
    openYoutubeBrowserWindow: () => youtube.browserWindow.open(),
    exportLogs: () => {
      void exportLogsFromTray();
    },
    openJellyfinSetupWindow: () => jellyfin.openJellyfinSetupWindow(),
    isJellyfinConfigured: () => isJellyfinConfiguredForTrayRuntime(jellyfin.getTrayDiscoveryDeps()),
    isJellyfinDiscoveryActive: () => Boolean(appState.jellyfinRemoteSession),
    toggleJellyfinDiscovery: (checked: boolean) =>
      toggleJellyfinDiscoveryFromTrayRuntime(jellyfin.getTrayDiscoveryDeps(), {
        desiredActive: checked,
      }),
    openAnilistSetupWindow: () => anilist.openSetupWindow(),
    checkForUpdates: () => {
      void getUpdateService().checkForUpdates({ source: 'manual' });
    },
    quitApp: () => requestAppQuit(),
  },
  initializeOverlayRuntime: () => startup.initializeOverlayRuntime(),
  setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
});
const overlay = createOverlayAssembly({
  appState,
  logger,
  overlayManager,
  overlayModalInputState,
  overlayModalRuntime,
  overlayContentMeasurementStore,
  getAutoplayReadyGate: () => playback.autoplayReadyGate,
  // Forward lazily: Electron's screen module must not be touched before app 'ready'.
  screen: {
    getCursorScreenPoint: () => screen.getCursorScreenPoint(),
    getDisplayNearestPoint: (point) => screen.getDisplayNearestPoint(point),
    getDisplayMatching: (rect) => screen.getDisplayMatching(rect),
    screenToDipPoint: (point) => screen.screenToDipPoint(point),
  },
  isDev,
  createOverlayWindowCore,
  getYomitanSession: () => yomitan.getOverlaySession(),
  getOverlayShortcuts: () => shortcuts.overlayShortcutsRuntime,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  getRuntimeOptionsState: () => getRuntimeOptionsStateHandler(),
  getOverlayForegroundSeparateWindows: () => getOverlayForegroundSeparateWindows(),
  promoteStatsOverlayAbovePlayback: () => {
    promoteStatsOverlayAbovePlayback();
  },
  getMediaTimingReviewRuntime: () => mining.mediaTimingReviewRuntime,
  getSubtitleProcessingController: () => subtitleProcessingController,
  getAutoplaySubtitlePriming: () => autoplaySubtitlePrimingRuntime,
  initializeOverlayRuntime: () => startup.initializeOverlayRuntime(),
  getPluginPanel: () => plugins.panel,
  setSubtitleSidebarRequestedOpen: (open) => playback.setSubtitleSidebarRequestedOpen(open),
  stealAppFocus: () => app.focus({ steal: true }),
});
const {
  overlayVisibilityRuntime,
  createMainWindow,
  createModalWindow,
  ensureOverlayWindowsReadyForVisibilityActions,
  setVisibleOverlayVisible,
  toggleVisibleOverlay,
  setOverlayVisible,
  handleMpvFullscreenChange,
  clearLinuxFullscreenOverlayRefresh,
  restoreOverlayMpvSubtitles,
  syncOverlayMpvSubtitleSuppression,
} = overlay;

const shortcuts = createShortcutsAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  configDir: CONFIG_DIR,
  isDev,
  overlayManager,
  globalShortcut,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  getOverlayModals: () => overlayModals,
  getMining: () => mining,
  triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
  cycleSecondarySubMode: () => sessionActions.cycleSecondarySubMode(),
  toggleVisibleOverlay: () => toggleVisibleOverlay(),
  openDictionarySettings: () => yomitan.openDictionarySettings(),
});

const settings = createSettingsAssembly({
  appState,
  configService,
  logger,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  hotReloadApplied: {
    setKeybindings: (keybindings) => {
      appState.keybindings = keybindings;
    },
    setSessionBindings: (sessionBindings, sessionBindingWarnings) => {
      shortcuts.persistSessionBindings(sessionBindings, sessionBindingWarnings);
    },
    refreshGlobalAndOverlayShortcuts: () => {
      shortcuts.refreshGlobalAndOverlayShortcuts();
    },
    setSecondarySubMode: (mode) => {
      sessionActions.setSecondarySubMode(mode);
    },
    broadcastToOverlayWindows: (channel, payload) => {
      overlayManager.broadcastToOverlayWindows(channel, payload);
    },
    applyAnkiRuntimeConfigPatch: (patch) => {
      if (appState.ankiIntegration) {
        appState.ankiIntegration.applyRuntimeConfigPatch(patch);
      }
      if (activeDictionaryBackend === 'hachidori' && appState.yomitanExt) {
        void yomitan
          .syncYomitanDefaultProfileAnkiServer()
          .catch((error: unknown) =>
            logger.error('Failed to auto-populate Hachidori Anki settings', error),
          );
      }
    },
    invalidateTokenizationCache: () => {
      subtitleProcessingController.invalidateTokenizationCache();
    },
    refreshSubtitlePrefetch: () => {
      getSubtitlePrefetchService()?.onSeek(getLastObservedTimePos());
    },
    refreshCurrentSubtitle: () => {
      subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
    },
    setLogLevel: (level) => {
      setLogLevel(level, 'config');
    },
    setLogRotation: (rotation) => {
      setLogRotation(rotation);
    },
    setLogFileToggles: (files) => {
      setLogFileToggles(files);
    },
    applyAniSkipConfig: () => {
      playback.aniSkipRuntime.applyConfigChange();
    },
  },
  getYomitanAnkiDeckName: () => yomitan.getCurrentAnkiDeckName(),
  showErrorBox: (title, content) => dialog.showErrorBox(title, content),
  createBrowserWindow: (options) => new BrowserWindow(options),
  mainDirname: __dirname,
  openPath: (targetPath) => shell.openPath(targetPath),
  ipcMain,
});

settings.window.registerHandlers();

const sync = createSyncAssembly({
  appState,
  configService,
  logger,
  userDataPath: USER_DATA_PATH,
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  requestAppQuit: () => requestAppQuit(),
  createBrowserWindow: (options) => new BrowserWindow(options),
  mainDirname: __dirname,
  ipcMain,
  isPackaged: () => app.isPackaged,
  getAppPath: () => app.getAppPath(),
  pickSnapshotFile: async () => {
    const result = await dialog.showOpenDialog({
      title: 'Choose a SubMiner stats snapshot',
      filters: [{ name: 'SQLite snapshots', extensions: ['sqlite'] }],
      properties: ['openFile'],
    });
    return result.canceled ? null : (result.filePaths[0] ?? null);
  },
  revealPath: (targetPath) => shell.showItemInFolder(targetPath),
});
sync.runtime.registerHandlers();
sync.autoScheduler.start();

const dictionaries = createDictionariesAssembly({
  roots: {
    platform: process.platform,
    dirname: __dirname,
    appPath: app.getAppPath(),
    resourcesPath: process.resourcesPath,
    userDataPath: USER_DATA_PATH,
    appUserDataPath: app.getPath('userData'),
    homeDir: os.homedir(),
    appDataDir: process.env.APPDATA,
    cwd: process.cwd(),
    joinPath: (...parts: string[]) => path.join(...parts),
  },
  appState,
  getConfig: () => configService.getConfig(),
  logger,
});

const mining = createMiningAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  overlayManager,
  overlayModalRuntime,
  youtube,
  clipboard,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  getOverlayModals: () => overlayModals,
  setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  ensureOverlayStartupPrereqs: () => startup.ensureOverlayStartupPrereqs(),
  ensureOverlayWindowsReadyForVisibilityActions: () =>
    ensureOverlayWindowsReadyForVisibilityActions(),
  ankiIntegrationCallbacks: {
    recordCardsMined: (count, noteIds) => stats.recordCardsMined(count, noteIds),
    getMiningProvenance: () => ({
      minedAt: Date.now(),
      media: plugins.host.getMediaSnapshot(),
      subtitle: plugins.getCurrentSubtitle(),
    }),
    onCardsReady: (noteIds, notes) => plugins.host.emitAnkiEvent('card-mined', { noteIds, notes }),
    onKnownWordCacheUpdated: () => refreshCurrentSubtitleAnnotations(),
    generateSentenceFurigana: (text, highlightedText) =>
      yomitan.generateSentenceFurigana(text, highlightedText),
  },
});

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, 'subtitle-positions');

const mediaRuntime = createMediaRuntimeService({
  isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath),
  loadSubtitlePosition: () => loadSubtitlePosition(),
  getCurrentMediaPath: () => appState.currentMediaPath,
  getPendingSubtitlePosition: () => appState.pendingSubtitlePosition,
  getSubtitlePositionsDir: () => SUBTITLE_POSITIONS_DIR,
  setCurrentMediaPath: (nextPath: string | null) => {
    appState.currentMediaPath = nextPath;
  },
  clearPendingSubtitlePosition: () => {
    appState.pendingSubtitlePosition = null;
  },
  setSubtitlePosition: (position: SubtitlePosition | null) => {
    appState.subtitlePosition = position;
  },
  broadcastSubtitlePosition: (position) => {
    overlayManager.broadcastToOverlayWindows('subtitle-position:set', position);
  },
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  setCurrentMediaTitle: (title) => {
    appState.currentMediaTitle = title;
  },
});

const characterDictionary = createCharacterDictionaryAssembly({
  userDataPath: USER_DATA_PATH,
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  startupOsdSequencer,
  resolveMediaPathForJimaku: (mediaPath) => mediaRuntime.resolveMediaPathForJimaku(mediaPath),
  getYomitanProfilePolicy: () => yomitan.profilePolicy,
  isYoutubePlaybackActive: () => youtube.isPlaybackActiveNow(),
  waitForMediaTokenizationReady: (mediaPath) =>
    playback.currentMediaTokenizationGate.waitUntilReady(mediaPath),
  ensureYomitanExtensionLoaded: () => yomitan.ensureExtensionLoaded(),
  getYomitanParserRuntimeDeps: () => yomitan.getParserRuntimeDeps(),
  getOverlayNotifications: () => overlayNotificationsRuntime,
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  invalidateTokenizationCache: () => subtitleProcessingController.invalidateTokenizationCache(),
  refreshCurrentSubtitle: (text) => subtitleProcessingController.refreshCurrentSubtitle(text),
  refreshSubtitlePrefetch: () => {
    getSubtitlePrefetchService()?.onSeek(getLastObservedTimePos());
  },
  refreshCurrentSubtitleAnnotations: () => refreshCurrentSubtitleAnnotations(),
  resetAnilistMediaGuessState: () => anilist.resetAnilistMediaGuessState(),
});

const getRuntimeOptionsStateHandler = createGetRuntimeOptionsStateHandler({
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
});

const overlayNotificationsRuntime = createOverlayNotificationsRuntime({
  getResolvedConfig: () => configService.getConfig(),
  getMainOverlayWindow: () => overlayManager.getMainWindow(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  broadcastToOverlayWindows: (channel, ...args) =>
    overlayManager.broadcastToOverlayWindows(channel, ...args),
  showMpvOsd: (message) => showMpvOsd(message),
  getMpvClient: () => appState.mpvClient,
  getAnkiIntegration: () => appState.ankiIntegration,
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
});
const {
  flushQueuedMpvOsdNotifications,
  openAnkiCardFromNotification,
  showConfiguredPlaybackFeedback,
  maybeStartOverlayLoadingOsd,
} = overlayNotificationsRuntime;

const overlayModals = createOverlayModalsAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  overlayModalRuntime,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  ensureOverlayStartupPrereqs: () => startup.ensureOverlayStartupPrereqs(),
  ensureOverlayWindowsReadyForVisibilityActions: () =>
    ensureOverlayWindowsReadyForVisibilityActions(),
  sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
    overlay.visibilityComposer.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  showMpvOsd: (text) => showMpvOsd(text),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
});

const plugins = createPluginsAssembly({
  appState,
  configService,
  overlayManager,
  logger,
  createLogger,
  configDir: CONFIG_DIR,
  userDataPath: USER_DATA_PATH,
  overlayModalRuntime,
  overlayModals,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  createBrowserWindow: (options) => new BrowserWindow(options),
  promoteAboveOverlay: (window) => promoteSettingsWindowAboveOverlay(window),
  ensureMpvRunning: () => jellyfin.ensureMpvConnectedForPlayback(),
  dispatchSessionAction: (request) => sessionActions.dispatchSessionAction(request),
});

function getRuntimeBooleanOption(
  id:
    | 'subtitle.annotation.knownWords.highlightEnabled'
    | 'subtitle.annotation.nPlusOne'
    | 'subtitle.annotation.jlpt'
    | 'subtitle.annotation.frequency',
  fallback: boolean,
): boolean {
  const value = appState.runtimeOptionsManager?.getOptionValue(id);
  return typeof value === 'boolean' ? value : fallback;
}

function shouldInitializeMecabForAnnotations(): boolean {
  const config = configService.getConfig();
  const knownWordsEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.knownWords.highlightEnabled',
    config.ankiConnect.knownWords.highlightEnabled,
  );
  const nPlusOneEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.nPlusOne',
    config.ankiConnect.nPlusOne.enabled,
  );
  const jlptEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.jlpt',
    config.subtitleStyle.enableJlpt,
  );
  const frequencyEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.frequency',
    config.subtitleStyle.frequencyDictionary.enabled,
  );
  return knownWordsEnabled || nPlusOneEnabled || jlptEnabled || frequencyEnabled;
}

const jellyfin = createJellyfinAssembly({
  appState,
  configService,
  jellyfinTokenStore,
  logger,
  defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
  getAppDataPath: () => app.getPath('appData'),
  getRuntimePluginEntrypoint: () => resolveBundledMpvRuntimePluginEntrypoint(),
  getMpvPluginRuntimeConfig: () => getMpvPluginRuntimeConfig(),
  createMpvClient: () => playback.createMpvClient(),
  initSubtitlePrefetch: (sourcePath) =>
    subtitlePrefetchRuntime.refreshSubtitleSidebarFromSource(sourcePath),
  showVisibleOverlay: () => setVisibleOverlayVisible(true),
  getOverlayNotifications: () => overlayNotificationsRuntime,
  updateCurrentMediaTitle: (title) => mediaRuntime.updateCurrentMediaTitle(title),
  ensureImmersionTrackerStarted: () => stats.ensureImmersionTrackerStarted(),
  refreshTrayMenu: () => tray.refreshMenu(),
  setupPreloadPath: path.join(__dirname, 'preload-jellyfin-setup.js'),
  createBrowserWindow: (options) => new BrowserWindow(options),
  ipcMain,
});

const anilist = createAnilistAssembly({
  appState,
  configService,
  logger,
  anilistTokenStore,
  anilistUpdateQueue,
  anilistRateLimiter,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  resolveMediaPathForJimaku: (currentMediaPath) =>
    mediaRuntime.resolveMediaPathForJimaku(currentMediaPath),
  resolvePinnedMediaId: (request) => characterDictionary.runtime.resolvePinnedMediaId(request),
  createBrowserWindow: (options) => new BrowserWindow(options),
  shell,
  app,
});

const loadSubtitlePosition = createLoadSubtitlePositionHandler({
  loadSubtitlePositionCore: () =>
    loadSubtitlePositionCore({
      currentMediaPath: appState.currentMediaPath,
      fallbackPosition: configService.getConfig().subtitlePosition,
      subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
    }),
  setSubtitlePosition: (position) => {
    appState.subtitlePosition = position;
  },
});

const saveSubtitlePosition = createSaveSubtitlePositionHandler({
  saveSubtitlePositionCore: (position) => {
    saveSubtitlePositionCore({
      position,
      currentMediaPath: appState.currentMediaPath,
      subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
      onQueuePending: (queued) => {
        appState.pendingSubtitlePosition = queued;
      },
      onPersisted: () => {
        appState.pendingSubtitlePosition = null;
      },
    });
  },
  setSubtitlePosition: (position) => {
    appState.subtitlePosition = position;
  },
});

anilist.registerSubminerProtocolClient();
let flushPendingMpvLogWrites = (): void => {};
const {
  registerProtocolUrlHandlers: registerProtocolUrlHandlersHandler,
  onWillQuitCleanup: onWillQuitCleanupHandler,
  shouldRestoreWindowsOnActivate: shouldRestoreWindowsOnActivateHandler,
  restoreWindowsOnActivate: restoreWindowsOnActivateHandler,
} = composeStartupLifecycleHandlers({
  registerProtocolUrlHandlersMainDeps: {
    registerOpenUrl: (listener) => {
      app.on('open-url', listener);
    },
    registerSecondInstance: (listener) => {
      registerSecondInstanceHandlerEarly(app, listener);
    },
    handleAnilistSetupProtocolUrl: (rawUrl) => anilist.handleAnilistSetupProtocolUrl(rawUrl),
    findAnilistSetupDeepLinkArgvUrl: (argv) => findAnilistSetupDeepLinkArgvUrl(argv),
    logUnhandledOpenUrl: (rawUrl) => {
      logger.warn('Unhandled app protocol URL', { rawUrl });
    },
    logUnhandledSecondInstanceUrl: (rawUrl) => {
      logger.warn('Unhandled second-instance protocol URL', { rawUrl });
    },
  },
  onWillQuitCleanupMainDeps: {
    destroyTray: () => tray.destroyTray(),
    stopConfigHotReload: () => settings.hotReload.stop(),
    restorePreviousSecondarySubVisibility: () =>
      overlay.visibilityComposer.restorePreviousSecondarySubVisibility(),
    restoreMpvSubVisibility: () => {
      restoreOverlayMpvSubtitles({ force: true });
    },
    isAppReady: () => app.isReady(),
    unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
    stopSubtitleWebsocket: () => {
      subtitleWsService.stop();
      annotationSubtitleWsService.stop();
    },
    stopTexthookerService: () => texthookerService.stop(),
    stopSyncAutoScheduler: async () => {
      sync.autoScheduler.stop();
      await sync.runtime.shutdown();
    },
    clearWindowsVisibleOverlayForegroundPollLoop: () =>
      overlay.interaction.clearWindowsVisibleOverlayForegroundPollLoop(),
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => clearLinuxFullscreenOverlayRefresh(),
    getMainOverlayWindow: () => overlayManager.getMainWindow(),
    clearMainOverlayWindow: () => overlayManager.setMainWindow(null),
    getModalOverlayWindow: () => overlayManager.getModalWindow(),
    clearModalOverlayWindow: () => overlayManager.setModalWindow(null),
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    clearYomitanParserState: () => {
      appState.yomitanParserWindow = null;
      appState.yomitanParserReadyPromise = null;
      appState.yomitanParserInitPromise = null;
      appState.yomitanSession = null;
    },
    getWindowTracker: () => appState.windowTracker,
    flushMpvLog: () => flushPendingMpvLogWrites(),
    getMpvSocket: () => appState.mpvClient?.socket ?? null,
    getReconnectTimer: () => appState.reconnectTimer,
    clearReconnectTimerRef: () => {
      appState.reconnectTimer = null;
    },
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getImmersionTracker: () => appState.immersionTracker,
    stopStatsServer: () => stats.stopStatsServer(),
    clearImmersionTracker: () => {
      appState.statsServer = null;
      appState.immersionTracker = null;
    },
    getAnkiIntegration: () => appState.ankiIntegration,
    getAnilistSetupWindow: () => appState.anilistSetupWindow,
    clearAnilistSetupWindow: () => {
      appState.anilistSetupWindow = null;
    },
    getJellyfinSetupWindow: () => appState.jellyfinSetupWindow,
    clearJellyfinSetupWindow: () => {
      appState.jellyfinSetupWindow = null;
    },
    getFirstRunSetupWindow: () => appState.firstRunSetupWindow,
    clearFirstRunSetupWindow: () => {
      appState.firstRunSetupWindow = null;
    },
    getYomitanSettingsWindow: () => appState.yomitanSettingsWindow,
    clearYomitanSettingsWindow: () => {
      appState.yomitanSettingsWindow = null;
    },
    stopJellyfinRemoteSession: () => jellyfin.stopJellyfinRemoteSession(),
    cleanupInternalSubtitleTrackCache: () => cachedInternalSubtitleTrackExtractor.clear(),
    cleanupYoutubeSubtitleTempDirs: () => {
      // Stops yt-dlp/whisper-cli and deletes the audio they were working on.
      youtube.flow.cancelWhisperGeneration();
      youtube.whisperSubtitles.removeActiveAudio();
      youtube.flow.cleanupSubtitleTempDirs();
    },
    cleanupYoutubeMediaCache: () => youtube.mediaCache.cleanup(),
    cleanupRemoteMediaWindows: () => getSharedRemoteMediaWindowCache().cleanup(),
    cleanupJellyfinSubtitleCache: () => jellyfin.cleanupJellyfinSubtitleCache(),
    stopDiscordPresenceService: () => {
      void appState.discordPresenceService?.stop();
      appState.discordPresenceService = null;
    },
    unloadPlugins: () => plugins.host.unloadAll(),
  },
  shouldRestoreWindowsOnActivateMainDeps: {
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    getAllWindowCount: () => BrowserWindow.getAllWindows().length,
  },
  restoreWindowsOnActivateMainDeps: {
    createMainWindow: () => {
      createMainWindow();
    },
    updateVisibleOverlayVisibility: () => {
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
    },
    syncOverlayMpvSubtitleSuppression: () => {
      syncOverlayMpvSubtitleSuppression();
    },
  },
});
registerProtocolUrlHandlersHandler();

const startup = createStartupAssembly({
  appState,
  configService,
  logger,
  logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
  configDir: CONFIG_DIR,
  userDataPath: USER_DATA_PATH,
  defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
  showErrorBox: (title, content) => dialog.showErrorBox(title, content),
  reportFatalError,
  requestAppQuit: () => requestAppQuit(),
  logging: {
    setLogLevel: (level: string, source: LogLevelSource) => setLogLevel(level, source),
    setLogRotation: (rotation: number) => setLogRotation(rotation),
    setLogFileToggles: (files) => setLogFileToggles(files),
  },
  overlayManager,
  overlayModalRuntime,
  subtitleWsService,
  annotationSubtitleWsService,
  texthookerService,
  configDerivedRuntime,
  hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
  loadSubtitlePosition: () => loadSubtitlePosition(),
  setSecondarySubMode: (mode) => sessionActions.setSecondarySubMode(mode),
  yomitan,
  stats,
  firstRunSetup,
  youtube,
  overlay,
  shortcuts,
  settings,
  mining,
  getAnilist: () => anilist,
  getJellyfin: () => jellyfin,
  getPlayback: () => playback,
  getPluginHost: () => plugins.host,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  handleInitialArgs: () => handleInitialArgs(),
});

const playback = createPlaybackAssembly({
  appState,
  configService,
  logger,
  overlayManager,
  subtitleWsService,
  annotationSubtitleWsService,
  overlayContentMeasurementStore,
  startupOsdSequencer,
  mediaTimingReviewRuntime: mining.mediaTimingReviewRuntime,
  discordPresenceRuntime,
  immersionMediaRuntime: stats.immersionMediaRuntime,
  mediaRuntime,
  overlayNotificationsRuntime,
  overlayModalRuntime,
  youtube,
  jellyfin,
  anilist,
  characterDictionary,
  dictionaries,
  getPluginHost: () => plugins.host,
  getSubtitleGenerationRuntime: () => subtitleGeneration.runtime,
  isCharacterDictionaryEnabled: () => yomitan.profilePolicy.isCharacterDictionaryEnabled(),
  hasTray: () => tray.hasTray(),
  loadSubtitleSourceText,
  createSubtitlePrefetchService,
  requestAppQuit: () => requestAppQuit(),
  maybeStartOverlayLoadingOsd: (mediaPath) => maybeStartOverlayLoadingOsd(mediaPath),
  flushQueuedMpvOsdNotifications: () => flushQueuedMpvOsdNotifications(),
  refreshMpvSessionBindings: () => shortcuts.refreshMpvSessionBindings(),
  ensureImmersionTrackerStarted: () => stats.ensureImmersionTrackerStarted(),
  restoreOverlayMpvSubtitles: () => restoreOverlayMpvSubtitles(),
  syncOverlayMpvSubtitleSuppression: () => syncOverlayMpvSubtitleSuppression(),
  setOverlayVisible: (visible) => setOverlayVisible(visible),
  ensureYomitanExtensionLoaded: () => yomitan.ensureExtensionLoaded(),
  shouldInitializeMecabForAnnotations: () => shouldInitializeMecabForAnnotations(),
  getRuntimeBooleanOption: (id, fallback) => getRuntimeBooleanOption(id, fallback),
  onFullscreenChange: (fullscreen) => handleMpvFullscreenChange(fullscreen),
  stealAppFocus: () => app.focus({ steal: true }),
});
const {
  subtitleProcessingController,
  autoplaySubtitlePrimingRuntime,
  subtitlePrefetchRuntime,
  cachedInternalSubtitleTrackExtractor,
  sendRendererMpvCommand,
  refreshCurrentSubtitleAnnotations,
  createMecabTokenizerAndCheck,
  startBackgroundWarmups,
  startTokenizationWarmups,
  getSubtitlePrefetchService,
  getLastObservedTimePos,
  getTokenizeSubtitleDeferred,
} = playback;

const { exportLogsFromTray } = createLogExportTrayRuntime({
  flushMpvLog: () => flushMpvLog(),
  logInfo: (message) => logger.info(message),
  logWarn: (message, details) => logger.warn(message, details),
});

const { flushMpvLog, showMpvOsd } = createMpvOsdRuntimeHandlers({
  appendToMpvLogMainDeps: {
    getLogPath: () => (isLogFileEnabled('mpv') ? DEFAULT_MPV_LOG_PATH : ''),
    dirname: (targetPath) => path.dirname(targetPath),
    mkdir: async (targetPath, options) => {
      await fs.promises.mkdir(targetPath, options);
    },
    appendFile: async (targetPath, data, options) => {
      await fs.promises.appendFile(targetPath, data, options);
    },
    now: () => new Date(),
  },
  buildShowMpvOsdMainDeps: (appendToMpvLogHandler) => ({
    appendToMpvLog: (message) => appendToMpvLogHandler(message),
    showMpvOsdRuntime: (mpvClient, text, fallbackLog) =>
      showMpvOsdRuntime(mpvClient, text, fallbackLog),
    getMpvClient: () => appState.mpvClient,
    logInfo: (line) => logger.info(line),
  }),
});
flushPendingMpvLogWrites = () => {
  void flushMpvLog();
};

const { getChangelogSnapshot } = createChangelogRuntime({
  getInstalledVersion: () => app.getVersion(),
  getUpdateChannel: () => configService.getConfig().updates.channel,
  resourcesPath: process.resourcesPath,
  appPath: app.getAppPath(),
  dirname: __dirname,
  joinPath: (...parts) => path.join(...parts),
  fileExists: (candidate) => fs.existsSync(candidate),
  readFile: (candidate) => fs.readFileSync(candidate, 'utf8'),
  logWarn: (message) => logger.warn(message),
});

const { getUpdateService, takePendingLauncherMigrationPath } = createUpdateServiceRuntime({
  userDataPath: USER_DATA_PATH,
  getUpdatesConfig: () => configService.getConfig().updates,
  logInfo: (message) => logger.info(message),
  logWarn: (message, details) => logger.warn(message, details),
  showOverlayNotification: (payload) =>
    overlayNotificationsRuntime.showOverlayNotification(payload),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  showMpvOsd: (message) => {
    showMpvOsd(message);
  },
  withStatsWindowLayerSuspendedForNativeDialog: (showDialog) =>
    withStatsWindowLayerSuspendedForNativeDialog(showDialog),
});

const sessionActions = createSessionActionsAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  overlayManager,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  stats,
  mining,
  overlayModals,
  youtube,
  plugins,
  toggleVisibleOverlay: () => toggleVisibleOverlay(),
  triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
  readClipboardText: () => clipboard.readText(),
  showMpvOsd: (text) => showMpvOsd(text),
  sendRendererMpvCommand: (command) => sendRendererMpvCommand(command),
});

const { playlistBrowserMainDeps } = createPlaylistBrowserIpcRuntime(() => appState.mpvClient, {
  getPrimarySubtitleLanguages: () => configService.getConfig().youtube.primarySubLanguages,
  getSecondarySubtitleLanguages: () => configService.getConfig().secondarySub.secondarySubLanguages,
});

const { registerIpcRuntimeHandlers } = composeIpcRuntimeHandlers({
  mpvCommandMainDeps: sessionActions.mpvCommandDeps,
  handleMpvCommandFromIpcRuntime,
  runSubsyncManualFromIpc: (request) => subsyncRuntime.runManualFromIpc(request),
  registration: {
    runtimeOptions: {
      getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
      showMpvOsd: (text: string) => showConfiguredPlaybackFeedback(text),
    },
    mainDeps: {
      ...overlay.ipc,
      ...playback.ipc,
      ...mining.ipc,
      ...shortcuts.ipc,
      ...anilist.ipc,
      ...characterDictionary.ipc,
      ...stats.ipc,
      ...playlistBrowserMainDeps,
      // App-level handlers with no owning assembly.
      onYoutubePickerResolve: (request) => youtube.flow.resolveActivePicker(request),
      handleOverlayNotificationAction: (notificationId, actionId, noteId) => {
        if (
          notificationId === UPDATE_AVAILABLE_NOTIFICATION_ID &&
          actionId === INSTALL_UPDATE_ACTION_ID
        ) {
          void getUpdateService()
            .checkForUpdates({
              source: 'manual',
              installWhenAvailable: true,
            })
            .catch((error) => {
              logger.warn('Failed to install update from overlay notification action:', error);
            });
        }
        if (
          notificationId === UPDATE_AVAILABLE_NOTIFICATION_ID &&
          actionId === VIEW_CHANGELOG_ACTION_ID
        ) {
          overlayModals.openChangelogOverlay();
        }
        if (actionId === OPEN_ANKI_CARD_ACTION_ID && noteId !== undefined) {
          void openAnkiCardFromNotification(noteId).catch((error) => {
            logger.warn('Failed to open Anki card from overlay notification action:', error);
            overlayNotificationsRuntime.showConfiguredStatusNotification(
              'Failed to open Anki card in Anki.',
              {
                id: 'open-anki-card-failed',
                variant: 'error',
              },
            );
          });
        }
      },
      dispatchSessionAction: (request) => sessionActions.dispatchSessionAction(request),
      getSubtitlePosition: () => loadSubtitlePosition(),
      saveSubtitlePosition: (position) => saveSubtitlePosition(position),
      getSubtitleStyle: () => {
        const resolvedConfig = configService.getConfig();
        return resolveSubtitleStyleForRenderer(resolvedConfig);
      },
      getStatsToggleKey: () => configService.getConfig().stats.toggleKey,
      getMarkWatchedKey: () => configService.getConfig().stats.markWatchedKey,
      getOverlayNotificationPosition: () => configService.getConfig().notifications.overlayPosition,
      getControllerConfig: () => configService.getConfig().controller,
      saveControllerConfig: (update) => {
        const currentRawConfig = configService.getRawConfig();
        configService.patchRawConfig({
          controller: applyControllerConfigUpdate(currentRawConfig.controller, update),
        });
      },
      getSecondarySubMode: () => appState.secondarySubMode,
      getRuntimeOptions: () => getRuntimeOptionsStateHandler(),
      getChangelogSnapshot: (options) => getChangelogSnapshot(options),
    } satisfies Required<IpcHandlerDeps>,
    ankiJimakuDeps: createAnkiJimakuIpcRuntimeServiceDeps({
      getResolvedConfig: () => configService.getConfig(),
      getMpvClient: () => appState.mpvClient,
      getAnkiIntegration: () => appState.ankiIntegration,
      getFieldGroupingResolver: () => mining.getFieldGroupingResolver(),
      setFieldGroupingResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) =>
        mining.setFieldGroupingResolver(resolver),
      parseMediaInfo: (mediaPath: string | null) =>
        parseMediaInfo(mediaRuntime.resolveMediaPathForJimaku(mediaPath)),
      getCurrentMediaPath: () => appState.currentMediaPath,
      jimakuFetchJson: <T>(
        endpoint: string,
        query?: Record<string, string | number | boolean | null | undefined>,
      ): Promise<JimakuApiResponse<T>> => configDerivedRuntime.jimakuFetchJson<T>(endpoint, query),
      getJimakuMaxEntryResults: () => configDerivedRuntime.getJimakuMaxEntryResults(),
      getJimakuLanguagePreference: () => configDerivedRuntime.getJimakuLanguagePreference(),
      resolveJimakuApiKey: () => configDerivedRuntime.resolveJimakuApiKey(),
      isRemoteMediaPath: (mediaPath: string) => isRemoteMediaPath(mediaPath),
      downloadToFile: (
        url: string,
        destPath: string,
        headers: Record<string, string>,
        downloadOptions?: { isAllowedRedirect?: (url: URL) => boolean },
      ) => downloadToFile(url, destPath, headers, downloadOptions),
    }),
    registerIpcRuntimeServices,
  },
});
const { handleCliCommand, handleInitialArgs } = composeCliStartupHandlers({
  cliCommandContextMainDeps: {
    ...mining.cli,
    ...shortcuts.cli,
    ...anilist.cli,
    ...characterDictionary.cli,
    ...stats.cli,
    // App-level commands with no owning assembly.
    appState,
    setLogLevel: (level) => setLogLevel(level, 'cli'),
    onMpvSocketPathChanged: (nextSocketPath, previousSocketPath) =>
      overlay.interaction.retargetOverlayWindowTrackerForMpvSocket(
        nextSocketPath,
        previousSocketPath,
      ),
    texthookerService,
    getResolvedConfig: () => configService.getConfig(),
    defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
    defaultAnnotationWebsocketPort: DEFAULT_CONFIG.annotationWebsocket.port,
    hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
    openExternal: (url: string) => shell.openExternal(url),
    logBrowserOpenError: (url: string, error: unknown) =>
      logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
    showMpvOsd: (text: string) =>
      overlayNotificationsRuntime.showConfiguredStatusNotification(text),
    showPlaybackFeedback: (text: string) => showConfiguredPlaybackFeedback(text),
    initializeOverlayRuntime: () => startup.initializeOverlayRuntime(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    togglePrimarySubtitleBar: () =>
      overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.primarySubtitleBarToggle),
    openFirstRunSetupWindow: (force?: boolean) => firstRunSetup.openWindow(force),
    setVisibleOverlayVisible: (visible: boolean) => setVisibleOverlayVisible(visible),
    triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
    openJellyfinSetupWindow: () => jellyfin.openJellyfinSetupWindow(),
    runJellyfinCommand: (argsFromCommand: CliArgs) => jellyfin.runJellyfinCommand(argsFromCommand),
    runUpdateCommand: async (argsFromCommand: CliArgs, source: CliCommandSource) => {
      await runUpdateCliCommand(argsFromCommand, source, {
        checkForUpdates: (request) => getUpdateService().checkForUpdates(request),
        writeResponse: (responsePath, payload) =>
          writeUpdateCliCommandResponse(responsePath, payload),
        logWarn: (message, error) => logger.warn(message, error),
      });
    },
    runEnsureLinuxRuntimePluginAssetsCommand: async (
      argsFromCommand: CliArgs,
      source: CliCommandSource,
    ) => {
      await runEnsureLinuxRuntimePluginAssetsCliCommand(
        argsFromCommand,
        {
          ensureLinuxRuntimePluginAssets: () => ensureLinuxRuntimePluginAssets(),
          writeResponse: (responsePath, payload) =>
            writeEnsureLinuxRuntimePluginAssetsCliCommandResponse(responsePath, payload),
          logWarn: (message, error) => logger.warn(message, error),
        },
        source,
      );
    },
    runYoutubePlaybackFlow: (request) => youtube.playback.runYoutubePlaybackFlow(request),
    openHachidoriSettings: () => yomitan.openHachidoriSettings(),
    openYomitanSettings: () => yomitan.openYomitanSettings(),
    openConfigSettingsWindow: () => settings.window.openWindow(),
    openSyncUiWindow: () => sync.openWindow(),
    openYoutubeBrowserWindow: () => youtube.browserWindow.open(),
    cycleSecondarySubMode: () => sessionActions.cycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => overlayModals.openRuntimeOptionsPalette(),
    printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
    stopApp: () => requestAppQuit(),
    hasMainWindow: () => Boolean(overlayManager.getMainWindow()),
    dispatchSessionAction: (request: SessionActionDispatchRequest) =>
      sessionActions.dispatchSessionAction(request),
    schedule: (fn: () => void, delayMs: number) => setTimeout(fn, delayMs),
    logInfo: (message: string) => logger.info(message),
    logDebug: (message: string) => logger.debug(message),
    logWarn: (message: string) => logger.warn(message),
    logError: (message: string, err: unknown) => logger.error(message, err),
  } satisfies Required<CliCommandContextMainDeps>,
  cliCommandRuntimeHandlerMainDeps: {
    handleTexthookerOnlyModeTransitionMainDeps: {
      isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
      ensureOverlayStartupPrereqs: () => startup.ensureOverlayStartupPrereqs(),
      setTexthookerOnlyMode: (enabled) => {
        appState.texthookerOnlyMode = enabled;
      },
      commandNeedsOverlayStartupPrereqs: (inputArgs) =>
        commandNeedsOverlayStartupPrereqs(inputArgs),
      startBackgroundWarmups: () => startBackgroundWarmups(),
      logInfo: (message: string) => logger.info(message),
    },
    ensureTrayForCommand: (args) => {
      if (args.background || args.managedPlayback) {
        tray.ensureTray();
      }
    },
    handleCliCommandRuntimeServiceWithContext: (args, source, cliContext) =>
      handleCliCommandRuntimeServiceWithContext(args, source, cliContext),
  },
  initialArgsRuntimeHandlerMainDeps: {
    getInitialArgs: () => appState.initialArgs,
    isBackgroundMode: () => appState.backgroundMode,
    shouldEnsureTrayOnStartup: () =>
      shouldEnsureTrayOnStartupForInitialArgs(process.platform, appState.initialArgs),
    shouldRunHeadlessInitialCommand: (args) => isHeadlessInitialCommand(args),
    ensureTray: () => tray.ensureTray(),
    isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
    hasImmersionTracker: () => Boolean(appState.immersionTracker),
    getMpvClient: () => appState.mpvClient,
    commandNeedsOverlayStartupPrereqs: (args) => commandNeedsOverlayStartupPrereqs(args),
    commandNeedsOverlayRuntime: (args) => commandNeedsOverlayRuntime(args),
    ensureOverlayStartupPrereqs: () => startup.ensureOverlayStartupPrereqs(),
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    initializeOverlayRuntime: () => startup.initializeOverlayRuntime(),
    logInfo: (message) => logger.info(message),
  },
});
const { runAndApplyStartupState } = composeHeadlessStartupHandlers<
  CliArgs,
  StartupState,
  ReturnType<typeof createStartupBootstrapRuntimeDeps>
>({
  startupRuntimeHandlersDeps: {
    appLifecycleRuntimeRunnerMainDeps: {
      app: appLifecycleApp,
      platform: process.platform,
      shouldStartApp: (nextArgs: CliArgs) => shouldStartApp(nextArgs),
      parseArgs: (argv: string[]) => parseArgs(argv),
      handleCliCommand: (nextArgs: CliArgs, source: CliCommandSource) =>
        handleCliCommand(nextArgs, source),
      printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
      logNoRunningInstance: () => appLogger.logNoRunningInstance(),
      startControlServer: (handleArgv: (argv: string[]) => void) => {
        const server = startAppControlServer({
          socketPath: getAppControlSocketPath({ configDir: CONFIG_DIR }),
          platform: process.platform,
          handleArgv,
          logDebug: (message) => logger.debug(message),
          logWarn: (message, error) => logger.warn(message, error),
        });
        return () => server.close();
      },
      onReady: () => startup.runAppReady(),
      onWillQuitCleanup: () => onWillQuitCleanupHandler(),
      shouldRestoreWindowsOnActivate: () => shouldRestoreWindowsOnActivateHandler(),
      restoreWindowsOnActivate: () => restoreWindowsOnActivateHandler(),
      shouldQuitOnWindowAllClosed: () =>
        shouldQuitOnWindowAllClosedForTrayState({
          backgroundMode: appState.backgroundMode,
          hasTray: tray.hasTray(),
        }),
    },
    createAppLifecycleRuntimeRunner: (params) => createAppLifecycleRuntimeRunner(params),
    buildStartupBootstrapMainDeps: (startAppLifecycle) => ({
      argv: process.argv,
      parseArgs: (argv: string[]) => parseArgs(argv),
      setLogLevel: (level: string, source: LogLevelSource) => {
        setLogLevel(level, source);
      },
      forceX11Backend: (args: CliArgs) => {
        forceX11Backend(args);
      },
      enforceUnsupportedWaylandMode: (args: CliArgs) => {
        enforceUnsupportedWaylandMode(args);
      },
      shouldStartApp: (args: CliArgs) => shouldStartApp(args),
      getDefaultSocketPath: () =>
        configService.getConfig().mpv.socketPath || getDefaultSocketPathHandler(),
      defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
      configDir: CONFIG_DIR,
      defaultConfig: DEFAULT_CONFIG,
      generateConfigTemplate: (config: ResolvedConfig) => generateConfigTemplate(config),
      generateDefaultConfigFile: (
        args: CliArgs,
        options: {
          configDir: string;
          defaultConfig: unknown;
          generateTemplate: (config: unknown) => string;
        },
      ) => generateDefaultConfigFile(args, options),
      setExitCode: (code) => {
        process.exitCode = code;
      },
      quitApp: () => requestAppQuit(),
      logGenerateConfigError: (message) => logger.error(message),
      startAppLifecycle,
    }),
    createStartupBootstrapRuntimeDeps: (deps) => createStartupBootstrapRuntimeDeps(deps),
    runStartupBootstrapRuntime,
    applyStartupState: (startupState) => applyStartupState(appState, startupState),
  },
});

runAndApplyStartupState();
void app.whenReady().then(() => {
  void takePendingLauncherMigrationPath(async (pendingLauncherPath) => {
    const acknowledgedPaths = await refreshManagedCommandLineLauncher({
      ...(await createCommandLineLauncherRuntimeOptions()),
      additionalLauncherPaths: pendingLauncherPath ? [pendingLauncherPath] : [],
    });
    return pendingLauncherPath !== undefined && acknowledgedPaths.includes(pendingLauncherPath);
  }).catch((error) => {
    logger.warn('Failed to refresh the installed command-line launcher', error);
  });
  if (!shouldStartAutomaticUpdateChecks(appState.initialArgs)) {
    return;
  }
  getUpdateService().startAutomaticChecks();
});
const startupModeFlags = getStartupModeFlags(appState.initialArgs);
const shouldUseMinimalStartup = startupModeFlags.shouldUseMinimalStartup;
const shouldSkipHeavyStartup = startupModeFlags.shouldSkipHeavyStartup;
if (!appState.initialArgs || (!shouldUseMinimalStartup && !shouldSkipHeavyStartup)) {
  if (isAnilistTrackingEnabled(configService.getConfig())) {
    void anilist
      .refreshClientSecretStateIfEnabled({
        force: true,
        allowSetupPrompt: false,
      })
      .catch((error) => {
        logger.error('Failed to refresh AniList client secret state during startup', error);
      });
    anilist.stateRuntime.refreshRetryQueueState();
  }
  void initializeDiscordPresenceService().catch((error) => {
    logger.error('Failed to initialize Discord presence service during startup', error);
  });
}
registerIpcRuntimeHandlers();
registerSubtitleSelectionIpc({
  ipc: ipcMain,
  isAllowedSender: (sender) =>
    [overlayManager.getMainWindow(), overlayManager.getModalWindow()].some(
      (window) => window && !window.isDestroyed() && window.webContents === sender,
    ),
  runtime: createSubtitleSelectionRuntime({
    isEnabled: () => configService.getConfig().subtitleSelection.enabled,
    getMpvClient: () => appState.mpvClient,
  }),
});
const subtitleGeneration = createSubtitleGenerationAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  getModelDirectory: getSubtitleGenerationModelDirectory,
  overlayManager,
  overlayModals,
  youtube,
  jellyfin,
  ipcMain,
});
subtitleGeneration.registerIpc();
app.on('before-quit', () => subtitleGeneration.runtime.cancel());
