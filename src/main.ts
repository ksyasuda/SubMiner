/*
  SubMiner - All-in-one sentence mining overlay
  Copyright (C) 2024 sudacode

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
  shell,
  protocol,
  Extension,
  Tray,
  dialog,
  screen,
} from 'electron';
import { applyControllerConfigUpdate } from './main/controller-config-update.js';

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
import { spawn } from 'node:child_process';
import * as os from 'os';
import * as path from 'path';
import type {
  JimakuApiResponse,
  KikuFieldGroupingChoice,
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
} from './types';
import { AnkiIntegration } from './anki-integration';
import { SubtitleTimingTracker } from './subtitle-timing-tracker';
import { isRemoteMediaPath } from './jimaku/utils';
import {
  createLogger,
  setLogLevel,
  resolveDefaultLogFilePath,
  type LogLevelSource,
} from './logger';
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
import { IPC_CHANNELS, type OverlayHostedModal } from './shared/ipc/contracts';
import { buildConfigParseErrorDetails, failStartupFromConfig } from './main/config-validation';
import {
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  createAnilistStateRuntime,
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createOpenAnilistSetupWindowHandler,
  findAnilistSetupDeepLinkArgvUrl,
  isAnilistTrackingEnabled,
  loadAnilistManualTokenEntry,
  openAnilistSetupInBrowser,
  rememberAnilistAttemptedUpdateKey,
} from './main/runtime/domains/anilist';
import { DEFAULT_MIN_WATCH_RATIO } from './shared/watch-threshold';
import { buildJellyfinSetupFormHtml } from './main/runtime/domains/jellyfin';
import {
  createBroadcastRuntimeOptionsChangedHandler,
  resolveSubtitleStyleForRenderer,
} from './main/runtime/domains/overlay';
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
  generateDefaultConfigFile,
  resolveConfiguredShortcuts,
  showDesktopNotification,
} from './core/utils';
import {
  ensureDefaultConfigBootstrap,
  getDefaultConfigFilePaths,
  resolveDefaultMpvInstallPaths,
} from './shared/setup-state';
import {
  MpvIpcClient,
  Texthooker,
  copyCurrentSubtitle as copyCurrentSubtitleCore,
  deleteYomitanDictionaryByTitle,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  hasMpvWebsocketPlugin,
  importYomitanDictionaryFromZip,
  loadYomitanExtension as loadYomitanExtensionCore,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  mineSentenceCard as mineSentenceCardCore,
  openYomitanSettingsWindow,
  playNextSubtitleRuntime,
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  replayCurrentSubtitleRuntime,
  runStartupBootstrapRuntime,
  clearYomitanParserCachesForWindow,
  sendMpvCommandRuntime,
  setMpvSubVisibilityRuntime,
  triggerFieldGrouping as triggerFieldGroupingCore,
  upsertYomitanDictionarySettings,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
} from './core/services';
import { shouldAutoOpenFirstRunSetup } from './main/first-run-runtime';
import { createWindowsMpvLaunchDeps, launchWindowsMpv } from './main/runtime/windows-mpv-launch';
import { shouldEnsureTrayOnStartupForInitialArgs } from './main/runtime/startup-tray-policy';
import { createImmersionTrackerStartupHandler } from './main/runtime/immersion-startup';
import { createPlaylistBrowserIpcRuntime } from './main/runtime/playlist-browser-ipc';
import { openPlaylistBrowser as openPlaylistBrowserRuntime } from './main/runtime/playlist-browser-open';
import {
  createRunStatsCliCommandHandler,
  writeStatsCliCommandResponse,
} from './main/runtime/stats-cli-command';
import {
  isBackgroundStatsServerProcessAlive,
  readBackgroundStatsServerState,
  removeBackgroundStatsServerState,
  resolveBackgroundStatsServerUrl,
  writeBackgroundStatsServerState,
} from './main/runtime/stats-daemon';
import { resolveLegacyVocabularyPosFromTokens } from './core/services/immersion-tracker/legacy-vocabulary-pos';
import { createAnilistUpdateQueue } from './core/services/anilist/anilist-update-queue';
import {
  guessAnilistMediaInfo,
  updateAnilistPostWatchProgress,
} from './core/services/anilist/anilist-updater';
import { createJellyfinTokenStore } from './core/services/jellyfin-token-store';
import { createAnilistTokenStore } from './core/services/anilist/anilist-token-store';
import {
  registerSecondInstanceHandlerEarly,
  requestSingleInstanceLockEarly,
  shouldBypassSingleInstanceLockForArgv,
} from './main/early-single-instance';
import { createIpcRuntimeBootstrap } from './main/ipc-runtime-bootstrap';
import { createMainBootRuntime } from './main/main-boot-runtime';
import { createMainEarlyRuntime } from './main/main-early-runtime';
import { createMainPlaybackRuntime } from './main/main-playback-runtime';
import { createDefaultSocketPathResolver } from './main/default-socket-path';
import {
  getRuntimeBooleanOption as getRuntimeBooleanOptionFromManager,
  shouldInitializeMecabForAnnotations as shouldInitializeMecabForAnnotationsFromRuntimeOptions,
} from './main/runtime-option-helpers';
import {
  getDefaultPasswordStore,
  getPasswordStoreArg,
  getStartupModeFlags,
  normalizePasswordStoreArg,
} from './main/startup-flags';
import { handleCliCommandRuntimeServiceWithContext } from './main/cli-runtime';
import { createAnilistRuntimeCoordinator } from './main/anilist-runtime-coordinator';
import { createJellyfinRuntimeCoordinator } from './main/jellyfin-runtime-coordinator';
import { createOverlayGeometryAccessors } from './main/overlay-geometry-accessors';
import { createOverlayUiBootstrapFromProcessState } from './main/overlay-ui-bootstrap-from-main-state';
import type { OverlayGeometryRuntime } from './main/overlay-geometry-runtime';
import { createOverlayModalRuntimeService } from './main/overlay-runtime';
import type { OverlayUiRuntime } from './main/overlay-ui-runtime';
import { createShortcutsRuntimeFromMainState } from './main/shortcuts-runtime';
import { createSubtitleDictionaryRuntimeCoordinator } from './main/subtitle-dictionary-runtime';
import { createYomitanRuntimeBootstrap } from './main/yomitan-runtime-bootstrap';
import { createOverlayModalInputState } from './main/runtime/overlay-modal-input-state';
import { openYoutubeTrackPicker } from './main/runtime/youtube-picker-open';
import { createMainStartupRuntimeFromProcessState } from './main/main-startup-runtime-coordinator';
import { createStartupLifecycleRuntime } from './main/startup-lifecycle-runtime';
import { type StatsRuntime } from './main/stats-runtime';
import { createStatsRuntimeFromMainState } from './main/stats-runtime-coordinator';
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from './main/frequency-dictionary-runtime';
import {
  createJlptDictionaryRuntimeService,
  getJlptDictionarySearchPaths,
} from './main/jlpt-runtime';
import { createCharacterDictionaryRuntimeService } from './main/character-dictionary-runtime';
import { createCharacterDictionaryAutoSyncRuntimeService } from './main/runtime/character-dictionary-auto-sync';
import { handleCharacterDictionaryAutoSyncComplete } from './main/runtime/character-dictionary-auto-sync-completion';
import { notifyCharacterDictionaryAutoSyncStatus } from './main/runtime/character-dictionary-auto-sync-notifications';
import { createCurrentMediaTokenizationGate } from './main/runtime/current-media-tokenization-gate';
import { createStartupOsdSequencer } from './main/runtime/startup-osd-sequencer';
import {
  createRefreshSubtitlePrefetchFromActiveTrackHandler,
  createResolveActiveSubtitleSidebarSourceHandler,
} from './main/runtime/subtitle-prefetch-runtime';
import { isYoutubePlaybackActive } from './main/runtime/youtube-playback';
import { shouldForceOverrideYomitanAnkiServer } from './main/runtime/yomitan-anki-server';
import { createStartupSequenceRuntime } from './main/startup-sequence-runtime';
import { type StartupState, applyStartupState, createAppState } from './main/state';
import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from './main/anilist-url-guard';
import {
  ConfigService,
  ConfigStartupParseError,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
} from './config';
import { resolveConfigDir } from './config/path-resolution';
import { parseSubtitleCues } from './core/services/subtitle-cue-parser';
import {
  createSubtitlePrefetchService,
  type SubtitlePrefetchService,
} from './core/services/subtitle-prefetch';
import {
  buildSubtitleSidebarSourceKey,
  resolveSubtitleSourcePath,
} from './main/runtime/subtitle-prefetch-source';
import { createSubtitlePrefetchInitController } from './main/runtime/subtitle-prefetch-init';
import { codecToExtension, getSubsyncConfig } from './subsync/utils';

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
const DEFAULT_MPV_LOG_FILE = resolveDefaultLogFilePath({
  platform: process.platform,
  homeDir: os.homedir(),
  appDataDir: process.env.APPDATA,
});
const ANILIST_SETUP_CLIENT_ID_URL = 'https://anilist.co/api/v2/oauth/authorize';
const ANILIST_SETUP_RESPONSE_TYPE = 'token';
const ANILIST_DEFAULT_CLIENT_ID = '36084';
const ANILIST_REDIRECT_URI = 'https://anilist.subminer.moe/';
const ANILIST_DEVELOPER_SETTINGS_URL = 'https://anilist.co/settings/developer';
const ANILIST_UPDATE_MIN_WATCH_SECONDS = 10 * 60;
const ANILIST_DURATION_RETRY_INTERVAL_MS = 15_000;
const ANILIST_MAX_ATTEMPTED_UPDATE_KEYS = 1000;
const TRAY_TOOLTIP = 'SubMiner';

const JELLYFIN_LANG_PREF = 'ja,jp,jpn,japanese,en,eng,english,enUS,en-US';
const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS = 3000;
const DISCORD_PRESENCE_APP_ID = '1475264834730856619';
const JELLYFIN_MPV_CONNECT_TIMEOUT_MS = 3000;
const JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS = 20000;
const YOUTUBE_MPV_CONNECT_TIMEOUT_MS = 3000;
const YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS = 10000;
const YOUTUBE_MPV_YTDL_FORMAT = 'bestvideo*+bestaudio/best';
const YOUTUBE_DIRECT_PLAYBACK_FORMAT = 'b';
const MPV_JELLYFIN_DEFAULT_ARGS = [
  '--sub-auto=fuzzy',
  '--sub-file-paths=.;subs;subtitles',
  '--sid=auto',
  '--secondary-sid=auto',
  '--secondary-sub-visibility=no',
  '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
  '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
] as const;

let yomitanLoadInFlight: Promise<Extension | null> | null = null;
let notifyAnilistTokenStoreWarning: (message: string) => void = () => {};

const isDev = process.argv.includes('--dev') || process.argv.includes('--debug');
const texthookerService = new Texthooker(() => {
  const config = getResolvedConfig();
  const characterDictionaryEnabled =
    config.anilist.characterDictionary.enabled &&
    yomitanProfilePolicy.isCharacterDictionaryEnabled();
  const knownAndNPlusOneEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.nPlusOne',
    config.ankiConnect.knownWords.highlightEnabled,
  );

  return {
    enableKnownWordColoring: knownAndNPlusOneEnabled,
    enableNPlusOneColoring: knownAndNPlusOneEnabled,
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
    knownWordColor: config.ankiConnect.knownWords.color,
    nPlusOneColor: config.ankiConnect.nPlusOne.nPlusOne,
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
let syncOverlayShortcutsForModal: (isActive: boolean) => void = () => {};
let syncOverlayVisibilityForModal: () => void = () => {};
let overlayUi: OverlayUiRuntime<BrowserWindow> | null = null;
let overlayGeometryRuntime: OverlayGeometryRuntime<BrowserWindow> | null = null;
const resolveDefaultSocketPath = createDefaultSocketPathResolver(process.platform);

function getDefaultSocketPath(): string {
  return resolveDefaultSocketPath();
}

const bootServices = createMainBootRuntime({
  platform: process.platform,
  argv: process.argv,
  appDataDir: process.env.APPDATA,
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  homeDir: os.homedir(),
  defaultMpvLogFile: DEFAULT_MPV_LOG_FILE,
  envMpvLog: process.env.SUBMINER_MPV_LOG,
  defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
  getDefaultSocketPath: () => getDefaultSocketPath(),
  app,
  dialog,
  overlay: {
    getSyncOverlayShortcutsForModal: () => syncOverlayShortcutsForModal,
    getSyncOverlayVisibilityForModal: () => syncOverlayVisibilityForModal,
    createModalWindow: () => overlayUi!.createModalWindow(),
    getOverlayGeometry: () => getCurrentOverlayGeometry(),
  },
  notifications: {
    notifyAnilistTokenStoreWarning: (message) => notifyAnilistTokenStoreWarning(message),
    requestAppQuit: () => requestAppQuit(),
  },
});
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

let forceQuitTimer: ReturnType<typeof setTimeout> | null = null;
let stats: StatsRuntime | null = null;

function requestAppQuit(): void {
  if (stats) {
    stats.cleanupBeforeQuit();
  }
  if (!forceQuitTimer) {
    forceQuitTimer = setTimeout(() => {
      logger.warn('App quit timed out; forcing process exit.');
      app.exit(0);
    }, 2000);
  }
  app.quit();
}

process.on('SIGINT', () => {
  requestAppQuit();
});
process.on('SIGTERM', () => {
  requestAppQuit();
});

const startBackgroundWarmupsIfAllowed = (): void => {
  mpvRuntime.startBackgroundWarmups();
};

const mainEarlyRuntime = createMainEarlyRuntime({
  platform: process.platform,
  configDir: CONFIG_DIR,
  homeDir: os.homedir(),
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  binaryPath: process.execPath,
  appPath: app.getAppPath(),
  resourcesPath: process.resourcesPath,
  appDataDir: app.getPath('appData'),
  desktopDir: app.getPath('desktop'),
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  defaultJimakuLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  defaultJimakuMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
  defaultJimakuApiBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
  jellyfinLangPref: JELLYFIN_LANG_PREF,
  youtube: {
    directPlaybackFormat: YOUTUBE_DIRECT_PLAYBACK_FORMAT,
    mpvYtdlFormat: YOUTUBE_MPV_YTDL_FORMAT,
    autoLaunchTimeoutMs: YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS,
    connectTimeoutMs: YOUTUBE_MPV_CONNECT_TIMEOUT_MS,
    logPath: DEFAULT_MPV_LOG_PATH,
  },
  discordPresenceAppId: DISCORD_PRESENCE_APP_ID,
  appState,
  getResolvedConfig: () => getResolvedConfig(),
  getFallbackDiscordMediaDurationSec: () =>
    anilist.getAnilistMediaGuessRuntimeState().mediaDurationSec,
  configService,
  overlay: {
    overlayManager,
    overlayModalRuntime,
    getOverlayUi: () => overlayUi,
    getOverlayGeometry: () => overlayGeometryRuntime!,
    ensureTray: () => {
      overlayUi?.ensureTray();
    },
    hasTray: () => Boolean(appTray),
  },
  yomitan: {
    ensureYomitanExtensionLoaded: () => yomitan.ensureYomitanExtensionLoaded(),
    getParserRuntimeDeps: () => yomitan.getParserRuntimeDeps(),
    openYomitanSettings: () => yomitan.openYomitanSettings(),
  },
  subtitle: {
    getSubtitle: () => subtitle,
  },
  tokenization: {
    startTokenizationWarmups: () => mpvRuntime.startTokenizationWarmups(),
    getGate: () => currentMediaTokenizationGate,
  },
  appReady: {
    ensureYoutubePlaybackRuntimeReady: () => ensureYoutubePlaybackRuntimeReady(),
  },
  shortcuts: {
    refreshGlobalAndOverlayShortcuts: () => shortcuts.refreshGlobalAndOverlayShortcuts(),
  },
  notifications: {
    showDesktopNotification,
    showErrorBox: (title, content) => dialog.showErrorBox(title, content),
  },
  mpv: {
    sendMpvCommandRuntime: (client, command) => sendMpvCommandRuntime(client, command),
    setSubVisibility: (visible) => setMpvSubVisibilityRuntime(appState.mpvClient, visible),
    showMpvOsd: (text) => mpvRuntime.showMpvOsd(text),
  },
  actions: {
    requestAppQuit,
    writeShortcutLink: (shortcutPath, operation, details) =>
      shell.writeShortcutLink(shortcutPath, operation, details),
  },
  logger,
});
const {
  firstRun,
  discordPresenceRuntime,
  initializeDiscordPresenceService,
  overlaySubtitleSuppression,
  startupSupport,
  youtube,
} = mainEarlyRuntime;
const {
  ensureOverlayMpvSubtitlesHidden,
  restoreOverlayMpvSubtitles,
  syncOverlayMpvSubtitleSuppression,
} = overlaySubtitleSuppression;
const { immersionMediaRuntime, configDerivedRuntime, subsyncRuntime, configHotReloadRuntime } =
  startupSupport;
const currentMediaTokenizationGate = createCurrentMediaTokenizationGate();
const startupOsdSequencer = createStartupOsdSequencer({
  showOsd: (message) => mpvRuntime.showMpvOsd(message),
});
const isYoutubePlaybackActiveNow = (): boolean =>
  isYoutubePlaybackActive(appState.currentMediaPath, appState.mpvClient?.currentVideoPath ?? null);

let appTray: Tray | null = null;

const subtitlePrefetchRuntime = {
  cancelPendingInit: () => subtitle.cancelPendingSubtitlePrefetchInit(),
  refreshSubtitleSidebarFromSource: (sourcePath: string) =>
    subtitle.refreshSubtitleSidebarFromSource(sourcePath),
  refreshSubtitlePrefetchFromActiveTrack: () => subtitle.refreshSubtitlePrefetchFromActiveTrack(),
  scheduleSubtitlePrefetchRefresh: (delayMs?: number) =>
    subtitle.scheduleSubtitlePrefetchRefresh(delayMs),
  clearScheduledSubtitlePrefetchRefresh: () => subtitle.clearScheduledSubtitlePrefetchRefresh(),
} as const;
const startupOverlayUiAdapter = {
  broadcastRuntimeOptionsChanged: () => overlayUi?.broadcastRuntimeOptionsChanged(),
  ensureOverlayWindowsReadyForVisibilityActions: () =>
    overlayUi?.ensureOverlayWindowsReadyForVisibilityActions(),
  ensureTray: () => overlayUi?.ensureTray(),
  initializeOverlayRuntime: () => overlayUi?.initializeOverlayRuntime(),
  openRuntimeOptionsPalette: () => overlayUi?.openRuntimeOptionsPalette(),
  setOverlayVisible: (visible: boolean) => overlayUi?.setOverlayVisible(visible),
  setVisibleOverlayVisible: (visible: boolean) => overlayUi?.setVisibleOverlayVisible(visible),
  toggleVisibleOverlay: () => overlayUi?.toggleVisibleOverlay(),
} as const;

const shortcutsBootstrap = createShortcutsRuntimeFromMainState({
  appState,
  getResolvedConfig: () => getResolvedConfig(),
  globalShortcut,
  registerGlobalShortcutsCore,
  isDev,
  overlay: {
    getOverlayUi: () => overlayUi,
    overlayManager,
    overlayModalRuntime,
  },
  actions: {
    showMpvOsd: (text) => mpvRuntime.showMpvOsd(text),
    openYomitanSettings: () => yomitan.openYomitanSettings(),
    triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
    handleCycleSecondarySubMode: () => mpvRuntime.cycleSecondarySubMode(),
    handleMultiCopyDigit: (count) => mining.handleMultiCopyDigit(count),
  },
  mining: {
    copyCurrentSubtitle: () => mining.copyCurrentSubtitle(),
    handleMineSentenceDigit: (count) => mining.handleMineSentenceDigit(count),
    markLastCardAsAudioCard: () => mining.markLastCardAsAudioCard(),
    mineSentenceCard: () => mining.mineSentenceCard(),
    triggerFieldGrouping: () => mining.triggerFieldGrouping(),
    updateLastCardFromClipboard: () => mining.updateLastCardFromClipboard(),
  },
});
const shortcuts = shortcutsBootstrap.shortcuts;
const overlayShortcutsRuntime = shortcutsBootstrap.overlayShortcutsRuntime;
syncOverlayShortcutsForModal = (isActive: boolean): void => {
  shortcutsBootstrap.syncOverlayShortcutsForModal(isActive);
};

const subtitleDictionaryOverlayUiAdapter = {
  setVisibleOverlayVisible: (visible: boolean) => overlayUi?.setVisibleOverlayVisible(visible),
  getRestoreVisibleOverlayOnModalClose: () =>
    overlayModalRuntime.getRestoreVisibleOverlayOnModalClose(),
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ) => overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
} as const;

const { subtitle, dictionarySupport } = createSubtitleDictionaryRuntimeCoordinator({
  env: {
    platform: process.platform,
    dirname: __dirname,
    appPath: app.getAppPath(),
    resourcesPath: process.resourcesPath,
    userDataPath: USER_DATA_PATH,
    appUserDataPath: app.getPath('userData'),
    homeDir: os.homedir(),
    appDataDir: process.env.APPDATA,
    cwd: process.cwd(),
    configDir: CONFIG_DIR,
    defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  },
  appState,
  getResolvedConfig: () => getResolvedConfig(),
  services: {
    subtitleWsService,
    annotationSubtitleWsService,
    overlayManager,
    startupOsdSequencer,
  },
  logging: {
    debug: (message, ...args) => logger.debug(message, ...args),
    info: (message, ...args) => logger.info(message, ...args),
    warn: (message, ...args) => logger.warn(message, ...args),
    error: (message, ...args) => logger.error(message, ...args),
  },
  subtitle: {
    parseSubtitleCues,
    createSubtitlePrefetchService: (deps) => createSubtitlePrefetchService(deps),
  },
  overlay: {
    getOverlayUi: () => subtitleDictionaryOverlayUiAdapter,
    showMpvOsd: (message) => mpvRuntime.showMpvOsd(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  },
  playback: {
    isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath),
    isYoutubePlaybackActive,
    waitForYomitanMutationReady: (mediaKey) =>
      currentMediaTokenizationGate.waitUntilReady(mediaKey),
  },
  anilist: {
    guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
  },
  yomitan: {
    isCharacterDictionaryEnabled: () => yomitanProfilePolicy.isCharacterDictionaryEnabled(),
    isExternalReadOnlyMode: () => yomitanProfilePolicy.isExternalReadOnlyMode(),
    logSkippedWrite: (message) => yomitanProfilePolicy.logSkippedWrite(message),
    ensureYomitanExtensionLoaded: () => yomitan.ensureYomitanExtensionLoaded(),
    getParserRuntimeDeps: () => yomitan.getParserRuntimeDeps(),
  },
});

const getResolvedConfig = () => configService.getConfig();

const getRuntimeBooleanOption = (
  id: 'subtitle.annotation.nPlusOne' | 'subtitle.annotation.jlpt' | 'subtitle.annotation.frequency',
  fallback: boolean,
): boolean =>
  getRuntimeBooleanOptionFromManager(
    (optionId) => appState.runtimeOptionsManager?.getOptionValue(optionId),
    id,
    fallback,
  );

const shouldInitializeMecabForAnnotations = (): boolean =>
  shouldInitializeMecabForAnnotationsFromRuntimeOptions({
    getResolvedConfig: () => getResolvedConfig(),
    getRuntimeBooleanOption: (id, fallback) => getRuntimeBooleanOption(id, fallback),
  });

const jellyfin = createJellyfinRuntimeCoordinator({
  getResolvedConfig: () => getResolvedConfig(),
  configService: {
    patchRawConfig: (patch) => {
      configService.patchRawConfig(patch as Parameters<typeof configService.patchRawConfig>[0]);
    },
  },
  tokenStore: jellyfinTokenStore,
  platform: process.platform,
  execPath: process.execPath,
  defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
  defaultMpvArgs: MPV_JELLYFIN_DEFAULT_ARGS,
  connectTimeoutMs: JELLYFIN_MPV_CONNECT_TIMEOUT_MS,
  autoLaunchTimeoutMs: JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS,
  langPref: JELLYFIN_LANG_PREF,
  progressIntervalMs: JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS,
  ticksPerSecond: JELLYFIN_TICKS_PER_SECOND,
  appState,
  actions: {
    createMpvClient: () => mpvRuntime.createMpvClientRuntimeService(),
    applyJellyfinMpvDefaults: (client) => startupSupport.applyJellyfinMpvDefaults(client),
    showMpvOsd: (message) => mpvRuntime.showMpvOsd(message),
  },
  logger,
});

const anilist = createAnilistRuntimeCoordinator({
  getResolvedConfig: () => getResolvedConfig(),
  isTrackingEnabled: (config) => isAnilistTrackingEnabled(config),
  tokenStore: anilistTokenStore,
  updateQueue: anilistUpdateQueue,
  appState,
  dictionarySupport,
  actions: {
    showMpvOsd: (message) => mpvRuntime.showMpvOsd(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  },
  logger,
  constants: {
    authorizeUrl: ANILIST_SETUP_CLIENT_ID_URL,
    clientId: ANILIST_DEFAULT_CLIENT_ID,
    responseType: ANILIST_SETUP_RESPONSE_TYPE,
    redirectUri: ANILIST_REDIRECT_URI,
    developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
    durationRetryIntervalMs: ANILIST_DURATION_RETRY_INTERVAL_MS,
    minWatchSeconds: ANILIST_UPDATE_MIN_WATCH_SECONDS,
    maxAttemptedUpdateKeys: ANILIST_MAX_ATTEMPTED_UPDATE_KEYS,
  },
});

anilist.registerSubminerProtocolClient();
let flushPendingMpvLogWrites = (): void => {};
const {
  registerProtocolUrlHandlers: registerProtocolUrlHandlersHandler,
  onWillQuitCleanup: onWillQuitCleanupHandler,
  shouldRestoreWindowsOnActivate: shouldRestoreWindowsOnActivateHandler,
  restoreWindowsOnActivate: restoreWindowsOnActivateHandler,
} = createStartupLifecycleRuntime({
  protocolUrl: {
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
  cleanup: {
    destroyTray: () => overlayUi?.destroyTray(),
    stopConfigHotReload: () => configHotReloadRuntime.stop(),
    restorePreviousSecondarySubVisibility: () => overlayUi?.restorePreviousSecondarySubVisibility(),
    restoreMpvSubVisibility: () => {
      restoreOverlayMpvSubtitles();
    },
    unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
    stopSubtitleWebsocket: () => {
      subtitleWsService.stop();
      annotationSubtitleWsService.stop();
    },
    stopTexthookerService: () => texthookerService.stop(),
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
    clearImmersionTracker: () => {
      stats?.stopStatsServer();
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
    stopDiscordPresenceService: () => {
      void appState.discordPresenceService?.stop();
      appState.discordPresenceService = null;
    },
  },
  shouldRestoreWindowsOnActivate: {
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    getAllWindowCount: () => BrowserWindow.getAllWindows().length,
  },
  restoreWindowsOnActivate: {
    createMainWindow: () => overlayUi!.createMainWindow(),
    updateVisibleOverlayVisibility: () => {
      overlayUi?.updateVisibleOverlayVisibility();
    },
    syncOverlayMpvSubtitleSuppression: () => {
      syncOverlayMpvSubtitleSuppression();
    },
  },
});
registerProtocolUrlHandlersHandler();

const statsCoordinator = createStatsRuntimeFromMainState({
  dirname: __dirname,
  userDataPath: USER_DATA_PATH,
  appState,
  getResolvedConfig: () => getResolvedConfig(),
  dictionarySupport: {
    getConfiguredDbPath: () => immersionMediaRuntime.getConfiguredDbPath(),
    seedImmersionMediaFromCurrentMedia: () => immersionMediaRuntime.seedFromCurrentMedia(),
  },
  overlay: {
    getOverlayGeometry: () => ({
      getCurrentOverlayGeometry: () => getCurrentOverlayGeometry(),
    }),
    updateVisibleOverlayVisibility: () => {
      overlayUi?.updateVisibleOverlayVisibility();
    },
  },
  mpvRuntime: {
    createMecabTokenizerAndCheck: () => mpvRuntime.createMecabTokenizerAndCheck(),
  },
  actions: {
    requestAppQuit,
  },
  logger,
});
const statsBootstrap = statsCoordinator.statsBootstrap;
stats = statsCoordinator.stats;
const ensureStatsServerStarted = (): string => statsCoordinator.ensureStatsServerStarted();
const ensureBackgroundStatsServerStarted = (): {
  url: string;
  runningInCurrentProcess: boolean;
} => statsCoordinator.ensureBackgroundStatsServerStarted();
const stopBackgroundStatsServer = async (): Promise<{ ok: boolean; stale: boolean }> =>
  await statsCoordinator.stopBackgroundStatsServer();
const ensureImmersionTrackerStarted = (): void => {
  statsCoordinator.ensureImmersionTrackerStarted();
};
const recordTrackedCardsMined = statsBootstrap.recordTrackedCardsMined;
const runStatsCliCommand = async (
  args: Pick<
    CliArgs,
    | 'statsResponsePath'
    | 'statsBackground'
    | 'statsStop'
    | 'statsCleanup'
    | 'statsCleanupVocab'
    | 'statsCleanupLifetime'
  >,
  source: CliCommandSource,
): Promise<void> => {
  await statsCoordinator.runStatsCliCommand(args, source);
};

const { mpvRuntime, mining } = createMainPlaybackRuntime({
  appState,
  logPath: DEFAULT_MPV_LOG_PATH,
  logger,
  getResolvedConfig: () => getResolvedConfig(),
  getRuntimeBooleanOption,
  subtitle,
  yomitan: {
    ensureYomitanExtensionLoaded: () => yomitan.ensureYomitanExtensionLoaded(),
    isCharacterDictionaryEnabled: () =>
      getResolvedConfig().anilist.characterDictionary.enabled &&
      yomitanProfilePolicy.isCharacterDictionaryEnabled() &&
      !isYoutubePlaybackActiveNow(),
  },
  currentMediaTokenizationGate,
  startupOsdSequencer,
  dictionarySupport,
  overlay: {
    broadcastToOverlayWindows: (channel, payload) => {
      overlayManager.broadcastToOverlayWindows(channel, payload);
    },
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getOverlayUi: () => startupOverlayUiAdapter,
  },
  lifecycle: {
    requestAppQuit,
    restoreOverlayMpvSubtitles,
    syncOverlayMpvSubtitleSuppression,
    publishDiscordPresence: () => {
      discordPresenceRuntime.publishDiscordPresence();
    },
  },
  stats: {
    ensureImmersionTrackerStarted: () => ensureImmersionTrackerStarted(),
  },
  anilist,
  jellyfin,
  youtube,
  mining: {
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker as SubtitleTimingTracker,
    getAnkiIntegration: () => appState.ankiIntegration,
    getMpvClient: () => appState.mpvClient,
    readClipboardText: () => clipboard.readText(),
    writeClipboardText: (text) => clipboard.writeText(text),
    updateLastCardFromClipboardCore,
    triggerFieldGroupingCore,
    markLastCardAsAudioCardCore,
    mineSentenceCardCore,
    handleMultiCopyDigitCore,
    copyCurrentSubtitleCore,
    handleMineSentenceDigitCore,
    getCurrentSecondarySubText: () => appState.mpvClient?.currentSecondarySubText || undefined,
    recordCardsMined: (count, noteIds) => {
      ensureImmersionTrackerStarted();
      appState.immersionTracker?.recordCardsMined(count, noteIds);
    },
  },
});
const playlistBrowserIpcRuntime = createPlaylistBrowserIpcRuntime(() => appState.mpvClient);

function createMpvClientRuntimeService(): MpvIpcClient {
  return mpvRuntime.createMpvClientRuntimeService();
}

function updateMpvSubtitleRenderMetrics(patch: Partial<MpvSubtitleRenderMetrics>): void {
  mpvRuntime.updateMpvSubtitleRenderMetrics(patch);
}

function isTokenizationWarmupReady(): boolean {
  return mpvRuntime.isTokenizationWarmupReady();
}

const overlayGeometryAccessors = createOverlayGeometryAccessors({
  getOverlayGeometryRuntime: () => overlayGeometryRuntime,
  getWindowTracker: () => appState.windowTracker,
  screen,
});
const { getOverlayGeometryFallback, getCurrentOverlayGeometry, geometryMatches } =
  overlayGeometryAccessors;

flushPendingMpvLogWrites = () => {
  void mpvRuntime.flushMpvLog();
};

const startupSequence = createStartupSequenceRuntime({
  appState: {
    initialArgs: appState.initialArgs,
    runtimeOptionsManager: appState.runtimeOptionsManager,
  },
  userDataPath: USER_DATA_PATH,
  getResolvedConfig: () => getResolvedConfig(),
  anilist,
  actions: {
    initializeDiscordPresenceService: async () => {
      await initializeDiscordPresenceService();
    },
    requestAppQuit,
  },
  logger,
});

let handleInitialArgsRef: (() => void) | null = null;

const startupRuntime = createMainStartupRuntimeFromProcessState<StartupState>({
  appState,
  appLifecycle: {
    app: appLifecycleApp,
    argv: process.argv,
    platform: process.platform,
  },
  config: {
    configService,
    configHotReloadRuntime,
    configDerivedRuntime,
    ensureDefaultConfigBootstrap: (options) => ensureDefaultConfigBootstrap(options as never),
    getDefaultConfigFilePaths,
    generateConfigTemplate,
    defaultConfig: DEFAULT_CONFIG,
    defaultKeybindings: DEFAULT_KEYBINDINGS,
    configDir: CONFIG_DIR,
  },
  logging: {
    appLogger,
    logger,
    setLogLevel,
  },
  shell: {
    dialog,
    shell,
    showDesktopNotification,
  },
  runtime: {
    subtitle,
    startupOverlayUiAdapter,
    overlayManager,
    firstRun,
    anilist,
    jellyfin,
    stats: {
      ensureImmersionTrackerStarted: () => ensureImmersionTrackerStarted(),
      runStatsCliCommand: (argsFromCommand, source) => runStatsCliCommand(argsFromCommand, source),
      immersion: statsBootstrap.immersion,
    },
    mining,
    texthookerService,
    yomitan: {
      loadYomitanExtension: () => yomitan.loadYomitanExtension(),
      ensureYomitanExtensionLoaded: () => yomitan.ensureYomitanExtensionLoaded(),
      openYomitanSettings: () => yomitan.openYomitanSettings(),
      getCharacterDictionaryDisabledReason: () =>
        yomitanProfilePolicy.getCharacterDictionaryDisabledReason(),
    },
    subsyncRuntime,
    dictionarySupport,
    subtitleWsService,
    annotationSubtitleWsService,
  },
  commands: {
    mpvRuntime,
    runHeadlessInitialCommand: async (): Promise<void> =>
      startupSequence.runHeadlessInitialCommand({
        handleInitialArgs: () => {
          if (!handleInitialArgsRef) {
            throw new Error('Initial args handler not initialized');
          }
          handleInitialArgsRef();
        },
      }),
    shortcuts,
    cycleSecondarySubMode: () => mpvRuntime.cycleSecondarySubMode(),
    hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
    showMpvOsd: (text) => mpvRuntime.showMpvOsd(text),
    shouldAutoOpenFirstRunSetup: (args) => shouldAutoOpenFirstRunSetup(args),
    youtube,
    shouldEnsureTrayOnStartupForInitialArgs: (platform, initialArgs) =>
      shouldEnsureTrayOnStartupForInitialArgs(platform, initialArgs ?? null),
    isHeadlessInitialCommand: (args) => isHeadlessInitialCommand(args),
    commandNeedsOverlayStartupPrereqs: (args) => commandNeedsOverlayStartupPrereqs(args),
    commandNeedsOverlayRuntime: (args) => commandNeedsOverlayRuntime(args),
    handleCliCommandRuntimeServiceWithContext: (
      args,
      source,
      cliContext: Parameters<typeof handleCliCommandRuntimeServiceWithContext>[2],
    ) => handleCliCommandRuntimeServiceWithContext(args, source, cliContext),
    shouldStartApp: (args) => shouldStartApp(args),
    parseArgs: (argv) => parseArgs(argv),
    printHelp,
    onWillQuitCleanupHandler: () => onWillQuitCleanupHandler(),
    shouldRestoreWindowsOnActivateHandler: () => shouldRestoreWindowsOnActivateHandler(),
    restoreWindowsOnActivateHandler: () => restoreWindowsOnActivateHandler(),
    forceX11Backend: (args) => forceX11Backend(args),
    enforceUnsupportedWaylandMode: (args) => enforceUnsupportedWaylandMode(args),
    getDefaultSocketPath: () => startupSupport.getDefaultSocketPath(),
    generateDefaultConfigFile,
    runStartupBootstrapRuntime: (deps) => runStartupBootstrapRuntime(deps),
    applyStartupState: (startupState) => applyStartupState(appState, startupState),
    getStartupModeFlags,
    requestAppQuit,
  },
  constants: {
    defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
  },
});

function ensureOverlayStartupPrereqs(): void {
  startupRuntime.appReady.ensureOverlayStartupPrereqs();
}

async function ensureYoutubePlaybackRuntimeReady(): Promise<void> {
  await startupRuntime.appReady.ensureYoutubePlaybackRuntimeReady();
}

const { registerIpcRuntimeHandlers } = createIpcRuntimeBootstrap({
  appState,
  userDataPath: USER_DATA_PATH,
  getResolvedConfig: () => getResolvedConfig(),
  configService,
  overlay: {
    manager: overlayManager,
    getOverlayUi: () => overlayUi ?? undefined,
    modalRuntime: overlayModalRuntime,
    contentMeasurementStore: overlayContentMeasurementStore,
  },
  subtitle,
  mpvRuntime,
  shortcuts,
  actions: {
    requestAppQuit,
    openYomitanSettings: () => yomitan.openYomitanSettings(),
    openPlaylistBrowser: () => {
      void openPlaylistBrowserRuntime({
        ensureOverlayStartupPrereqs: () => startupRuntime.appReady.ensureOverlayStartupPrereqs(),
        ensureOverlayWindowsReadyForVisibilityActions: () =>
          overlayUi?.ensureOverlayWindowsReadyForVisibilityActions(),
        sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
          overlayUi?.sendToActiveOverlayWindow(channel, payload, runtimeOptions) ?? false,
      });
    },
    showDesktopNotification,
    setAnkiIntegration: (integration: AnkiIntegration | null) => {
      appState.ankiIntegration = integration;
      appState.ankiIntegration?.setRecordCardsMinedCallback(recordTrackedCardsMined);
    },
  },
  runtimes: {
    youtube,
    anilist,
    mining,
    dictionarySupport,
    playlistBrowser: {
      playlistBrowserMainDeps: playlistBrowserIpcRuntime.playlistBrowserMainDeps,
    },
    configDerived: configDerivedRuntime,
    subsync: subsyncRuntime,
  },
});
const { handleCliCommand, handleInitialArgs, runAndApplyStartupState } = startupRuntime;
handleInitialArgsRef = handleInitialArgs;

runAndApplyStartupState();
startupSequence.runPostStartupInitialization();
const { yomitan, yomitanProfilePolicy } = createYomitanRuntimeBootstrap({
  userDataPath: USER_DATA_PATH,
  getResolvedConfig: () => getResolvedConfig(),
  appState,
  loadYomitanExtensionCore,
  getLoadInFlight: () => yomitanLoadInFlight,
  setLoadInFlight: (promise) => {
    yomitanLoadInFlight = promise;
  },
  openYomitanSettingsWindow,
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
  logError: (message, error) => logger.error(message, error),
  showMpvOsd: (message) => mpvRuntime.showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
});
const overlayUiBootstrap = createOverlayUiBootstrapFromProcessState<BrowserWindow>({
  appState,
  overlayManager,
  overlayModalInputState,
  overlayModalRuntime,
  overlayShortcutsRuntime,
  runtimes: {
    dictionarySupport,
    firstRun,
    yomitan: {
      openYomitanSettings: () => yomitan.openYomitanSettings(),
    },
    jellyfin,
    anilist,
    shortcuts,
    mpvRuntime,
  },
  env: {
    screen,
    appPath: app.getAppPath(),
    resourcesPath: process.resourcesPath,
    dirname: __dirname,
    platform: process.platform,
    isDev,
  },
  actions: {
    showMpvOsd: (message) => mpvRuntime.showMpvOsd(message),
    showDesktopNotification,
    sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
    ensureOverlayMpvSubtitlesHidden,
    syncOverlayMpvSubtitleSuppression: () => syncOverlayMpvSubtitleSuppression(),
    getResolvedConfig: () => getResolvedConfig(),
    requestAppQuit,
  },
  trayState: {
    getTray: () => appTray,
    setTray: (tray) => {
      appTray = tray as Tray | null;
    },
    trayTooltip: TRAY_TOOLTIP,
    logWarn: (message) => logger.warn(message),
  },
  startup: {
    shouldSkipHeadlessOverlayBootstrap: () =>
      Boolean(appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs)),
    getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
    onInitialized: () => {
      appState.ankiIntegration?.setRecordCardsMinedCallback(recordTrackedCardsMined);
    },
  },
});
overlayGeometryRuntime = overlayUiBootstrap.overlayGeometry;
overlayUi = overlayUiBootstrap.overlayUi;
syncOverlayVisibilityForModal = overlayUiBootstrap.syncOverlayVisibilityForModal;

registerIpcRuntimeHandlers();
