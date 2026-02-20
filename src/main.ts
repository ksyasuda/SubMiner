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
  globalShortcut,
  clipboard,
  shell,
  protocol,
  Extension,
  Menu,
  Tray,
  nativeImage,
  dialog,
} from 'electron';

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

import * as path from 'path';
import * as os from 'os';
import * as fs from 'fs';
import { spawn } from 'node:child_process';
import { MecabTokenizer } from './mecab-tokenizer';
import type {
  JimakuApiResponse,
  SubtitleData,
  SubtitlePosition,
  WindowGeometry,
  SecondarySubMode,
  SubsyncManualRunRequest,
  SubsyncResult,
  KikuFieldGroupingChoice,
  RuntimeOptionState,
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
} from './types';
import { SubtitleTimingTracker } from './subtitle-timing-tracker';
import { AnkiIntegration } from './anki-integration';
import { RuntimeOptionsManager } from './runtime-options';
import { downloadToFile, isRemoteMediaPath, parseMediaInfo } from './jimaku/utils';
import { createLogger, setLogLevel, type LogLevelSource } from './logger';
import { commandNeedsOverlayRuntime, parseArgs, shouldStartApp } from './cli/args';
import type { CliArgs, CliCommandSource } from './cli/args';
import { printHelp } from './cli/help';
import {
  createCriticalConfigErrorHandler,
  createReloadConfigHandler,
} from './main/runtime/startup-config';
import { buildConfigWarningNotificationBody } from './main/config-validation';
import { createImmersionTrackerStartupHandler } from './main/runtime/immersion-startup';
import { createImmersionMediaRuntime } from './main/runtime/immersion-media';
import { createAnilistStateRuntime } from './main/runtime/anilist-state';
import { createConfigDerivedRuntime } from './main/runtime/config-derived';
import { appendClipboardVideoToQueueRuntime } from './main/runtime/clipboard-queue';
import { createMainSubsyncRuntime } from './main/runtime/subsync-runtime';
import {
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  findAnilistSetupDeepLinkArgvUrl,
  isAnilistTrackingEnabled,
  loadAnilistManualTokenEntry,
  loadAnilistSetupFallback,
  openAnilistSetupInBrowser,
} from './main/runtime/anilist-setup';
import {
  createConsumeAnilistSetupTokenFromUrlHandler,
  createHandleAnilistSetupProtocolUrlHandler,
  createNotifyAnilistSetupHandler,
  createRegisterSubminerProtocolClientHandler,
} from './main/runtime/anilist-setup-protocol';
import { createRefreshAnilistClientSecretStateHandler } from './main/runtime/anilist-token-refresh';
import {
  createHandleJellyfinRemoteGeneralCommand,
  createHandleJellyfinRemotePlay,
  createHandleJellyfinRemotePlaystate,
  getConfiguredJellyfinSession,
  type ActiveJellyfinRemotePlaybackState,
} from './main/runtime/jellyfin-remote-commands';
import {
  createReportJellyfinRemoteProgressHandler,
  createReportJellyfinRemoteStoppedHandler,
} from './main/runtime/jellyfin-remote-playback';
import {
  createEnsureMpvConnectedForJellyfinPlaybackHandler,
  createLaunchMpvIdleForJellyfinPlaybackHandler,
  createWaitForMpvConnectedHandler,
} from './main/runtime/jellyfin-remote-connection';
import {
  createHandleJellyfinSetupWindowClosedHandler,
  buildJellyfinSetupFormHtml,
  createHandleJellyfinSetupNavigationHandler,
  createHandleJellyfinSetupWindowOpenedHandler,
  createHandleJellyfinSetupSubmissionHandler,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
  parseJellyfinSetupSubmissionUrl,
} from './main/runtime/jellyfin-setup-window';
import {
  createHandleAnilistSetupWindowClosedHandler,
  createHandleAnilistSetupWindowOpenedHandler,
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createAnilistSetupDidFailLoadHandler,
  createAnilistSetupDidFinishLoadHandler,
  createAnilistSetupDidNavigateHandler,
  createAnilistSetupFallbackHandler,
  createAnilistSetupWillNavigateHandler,
  createAnilistSetupWillRedirectHandler,
  createAnilistSetupWindowOpenHandler,
  createHandleManualAnilistSetupSubmissionHandler,
} from './main/runtime/anilist-setup-window';
import {
  createEnsureAnilistMediaGuessHandler,
  createMaybeProbeAnilistDurationHandler,
} from './main/runtime/anilist-media-guess';
import {
  buildAnilistAttemptKey,
  createMaybeRunAnilistPostWatchUpdateHandler,
  createProcessNextAnilistRetryUpdateHandler,
  rememberAnilistAttemptedUpdateKey,
} from './main/runtime/anilist-post-watch';
import {
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
} from './main/runtime/subtitle-position';
import { registerProtocolUrlHandlers } from './main/runtime/protocol-url-handlers';
import { createHandleJellyfinAuthCommands } from './main/runtime/jellyfin-cli-auth';
import { createHandleJellyfinListCommands } from './main/runtime/jellyfin-cli-list';
import { createHandleJellyfinPlayCommand } from './main/runtime/jellyfin-cli-play';
import { createHandleJellyfinRemoteAnnounceCommand } from './main/runtime/jellyfin-cli-remote-announce';
import {
  createStartJellyfinRemoteSessionHandler,
  createStopJellyfinRemoteSessionHandler,
} from './main/runtime/jellyfin-remote-session-lifecycle';
import { createHandleInitialArgsHandler } from './main/runtime/initial-args-handler';
import { createHandleTexthookerOnlyModeTransitionHandler } from './main/runtime/cli-command-prechecks';
import { createCliCommandContext } from './main/runtime/cli-command-context';
import {
  createBindMpvClientEventHandlers,
  createHandleMpvConnectionChangeHandler,
  createHandleMpvSubtitleTimingHandler,
} from './main/runtime/mpv-client-event-bindings';
import { createMpvClientRuntimeServiceFactory } from './main/runtime/mpv-client-runtime-service';
import { createUpdateMpvSubtitleRenderMetricsHandler } from './main/runtime/mpv-subtitle-render-metrics';
import {
  createLaunchBackgroundWarmupTaskHandler,
  createStartBackgroundWarmupsHandler,
} from './main/runtime/startup-warmups';
import {
  buildRestartRequiredConfigMessage,
  createConfigHotReloadAppliedHandler,
  createConfigHotReloadMessageHandler,
  resolveSubtitleStyleForRenderer,
} from './main/runtime/config-hot-reload-handlers';
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
  generateDefaultConfigFile,
  resolveConfiguredShortcuts,
  resolveKeybindings,
  showDesktopNotification,
} from './core/utils';
import {
  MpvIpcClient,
  SubtitleWebSocket,
  Texthooker,
  applyMpvSubtitleRenderMetricsPatch,
  broadcastRuntimeOptionsChangedRuntime,
  copyCurrentSubtitle as copyCurrentSubtitleCore,
  createOverlayManager,
  createFieldGroupingOverlayRuntime,
  createNumericShortcutRuntime,
  createOverlayContentMeasurementStore,
  createSubtitleProcessingController,
  createOverlayWindow as createOverlayWindowCore,
  createTokenizerDepsRuntime,
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  enforceOverlayLayerOrder as enforceOverlayLayerOrderCore,
  ensureOverlayWindowLevel as ensureOverlayWindowLevelCore,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  hasMpvWebsocketPlugin,
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  loadSubtitlePosition as loadSubtitlePositionCore,
  loadYomitanExtension as loadYomitanExtensionCore,
  listJellyfinItemsRuntime,
  listJellyfinLibrariesRuntime,
  listJellyfinSubtitleTracksRuntime,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
  ImmersionTrackerService,
  JellyfinRemoteSessionService,
  mineSentenceCard as mineSentenceCardCore,
  openYomitanSettingsWindow,
  playNextSubtitleRuntime,
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  replayCurrentSubtitleRuntime,
  runStartupBootstrapRuntime,
  saveSubtitlePosition as saveSubtitlePositionCore,
  authenticateWithPasswordRuntime,
  createConfigHotReloadRuntime,
  resolveJellyfinPlaybackPlanRuntime,
  jellyfinTicksToSecondsRuntime,
  sendMpvCommandRuntime,
  setInvisibleOverlayVisible as setInvisibleOverlayVisibleCore,
  setMpvSubVisibilityRuntime,
  setOverlayDebugVisualizationEnabledRuntime,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  showMpvOsdRuntime,
  tokenizeSubtitle as tokenizeSubtitleCore,
  triggerFieldGrouping as triggerFieldGroupingCore,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
} from './core/services';
import {
  guessAnilistMediaInfo,
  type AnilistMediaGuess,
  updateAnilistPostWatchProgress,
} from './core/services/anilist/anilist-updater';
import { createAnilistTokenStore } from './core/services/anilist/anilist-token-store';
import { createAnilistUpdateQueue } from './core/services/anilist/anilist-update-queue';
import { applyRuntimeOptionResultRuntime } from './core/services/runtime-options-ipc';
import { createAppReadyRuntimeRunner } from './main/app-lifecycle';
import { handleMpvCommandFromIpcRuntime } from './main/ipc-mpv-command';
import { registerIpcRuntimeServices } from './main/ipc-runtime';
import { createAnkiJimakuIpcRuntimeServiceDeps } from './main/dependencies';
import { handleCliCommandRuntimeServiceWithContext } from './main/cli-runtime';
import { createOverlayModalRuntimeService, type OverlayHostedModal } from './main/overlay-runtime';
import { createOverlayShortcutsRuntimeService } from './main/overlay-shortcuts-runtime';
import {
  createJlptDictionaryRuntimeService,
  getJlptDictionarySearchPaths,
} from './main/jlpt-runtime';
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from './main/frequency-dictionary-runtime';
import { createMediaRuntimeService } from './main/media-runtime';
import { createOverlayVisibilityRuntimeService } from './main/overlay-visibility-runtime';
import { type AppState, applyStartupState, createAppState } from './main/state';
import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from './main/anilist-url-guard';
import { createStartupBootstrapRuntimeDeps } from './main/startup';
import { createAppLifecycleRuntimeRunner } from './main/startup-lifecycle';
import {
  ConfigService,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
} from './config';
import { resolveConfigDir } from './config/path-resolution';

if (process.platform === 'linux') {
  app.commandLine.appendSwitch('enable-features', 'GlobalShortcutsPortal');
}

app.setName('SubMiner');

const DEFAULT_TEXTHOOKER_PORT = 5174;
const DEFAULT_MPV_LOG_FILE = path.join(os.homedir(), '.cache', 'SubMiner', 'mp.log');
const ANILIST_SETUP_CLIENT_ID_URL = 'https://anilist.co/api/v2/oauth/authorize';
const ANILIST_SETUP_RESPONSE_TYPE = 'token';
const ANILIST_DEFAULT_CLIENT_ID = '36084';
const ANILIST_REDIRECT_URI = 'https://anilist.subminer.moe/';
const ANILIST_DEVELOPER_SETTINGS_URL = 'https://anilist.co/settings/developer';
const ANILIST_UPDATE_MIN_WATCH_RATIO = 0.85;
const ANILIST_UPDATE_MIN_WATCH_SECONDS = 10 * 60;
const ANILIST_DURATION_RETRY_INTERVAL_MS = 15_000;
const ANILIST_MAX_ATTEMPTED_UPDATE_KEYS = 1000;
const ANILIST_TOKEN_STORE_FILE = 'anilist-token-store.json';
const ANILIST_RETRY_QUEUE_FILE = 'anilist-retry-queue.json';
const TRAY_TOOLTIP = 'SubMiner';

let anilistCurrentMediaKey: string | null = null;
let anilistCurrentMediaDurationSec: number | null = null;
let anilistCurrentMediaGuess: AnilistMediaGuess | null = null;
let anilistCurrentMediaGuessPromise: Promise<AnilistMediaGuess | null> | null = null;
let anilistLastDurationProbeAtMs = 0;
let anilistUpdateInFlight = false;
const anilistAttemptedUpdateKeys = new Set<string>();
let anilistCachedAccessToken: string | null = null;
let jellyfinPlayQuitOnDisconnectArmed = false;
const JELLYFIN_LANG_PREF = 'ja,jp,jpn,japanese,en,eng,english,enUS,en-US';
const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS = 3000;
const JELLYFIN_MPV_CONNECT_TIMEOUT_MS = 3000;
const JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS = 20000;
const MPV_JELLYFIN_DEFAULT_ARGS = [
  '--sub-auto=fuzzy',
  '--sub-file-paths=.;subs;subtitles',
  '--sid=auto',
  '--secondary-sid=auto',
  '--secondary-sub-visibility=no',
  '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
  '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
] as const;

let activeJellyfinRemotePlayback: ActiveJellyfinRemotePlaybackState | null = null;
let jellyfinRemoteLastProgressAtMs = 0;
let jellyfinMpvAutoLaunchInFlight: Promise<boolean> | null = null;
let backgroundWarmupsStarted = false;
let yomitanLoadInFlight: Promise<Extension | null> | null = null;

function applyJellyfinMpvDefaults(client: MpvIpcClient): void {
  sendMpvCommandRuntime(client, ['set_property', 'sub-auto', 'fuzzy']);
  sendMpvCommandRuntime(client, ['set_property', 'aid', 'auto']);
  sendMpvCommandRuntime(client, ['set_property', 'sid', 'auto']);
  sendMpvCommandRuntime(client, ['set_property', 'secondary-sid', 'auto']);
  sendMpvCommandRuntime(client, ['set_property', 'secondary-sub-visibility', 'no']);
  sendMpvCommandRuntime(client, ['set_property', 'alang', JELLYFIN_LANG_PREF]);
  sendMpvCommandRuntime(client, ['set_property', 'slang', JELLYFIN_LANG_PREF]);
}

const CONFIG_DIR = resolveConfigDir({
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  homeDir: os.homedir(),
  existsSync: fs.existsSync,
});
const USER_DATA_PATH = CONFIG_DIR;
const DEFAULT_MPV_LOG_PATH = process.env.SUBMINER_MPV_LOG?.trim() || DEFAULT_MPV_LOG_FILE;
const DEFAULT_IMMERSION_DB_PATH = path.join(USER_DATA_PATH, 'immersion.sqlite');
const configService = new ConfigService(CONFIG_DIR);
const anilistTokenStore = createAnilistTokenStore(
  path.join(USER_DATA_PATH, ANILIST_TOKEN_STORE_FILE),
  {
    info: (message: string) => console.info(message),
    warn: (message: string, details?: unknown) => console.warn(message, details),
    error: (message: string, details?: unknown) => console.error(message, details),
  },
);
const anilistUpdateQueue = createAnilistUpdateQueue(
  path.join(USER_DATA_PATH, ANILIST_RETRY_QUEUE_FILE),
  {
    info: (message: string) => console.info(message),
    warn: (message: string, details?: unknown) => console.warn(message, details),
    error: (message: string, details?: unknown) => console.error(message, details),
  },
);
const isDev = process.argv.includes('--dev') || process.argv.includes('--debug');
const texthookerService = new Texthooker();
const subtitleWsService = new SubtitleWebSocket();
const logger = createLogger('main');
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

function getDefaultSocketPath(): string {
  if (process.platform === 'win32') {
    return '\\\\.\\pipe\\subminer-socket';
  }
  return '/tmp/subminer-socket';
}

if (!fs.existsSync(USER_DATA_PATH)) {
  fs.mkdirSync(USER_DATA_PATH, { recursive: true });
}
app.setPath('userData', USER_DATA_PATH);

process.on('SIGINT', () => {
  app.quit();
});
process.on('SIGTERM', () => {
  app.quit();
});

const overlayManager = createOverlayManager();
const overlayContentMeasurementStore = createOverlayContentMeasurementStore({
  now: () => Date.now(),
  warn: (message: string) => logger.warn(message),
});
const overlayModalRuntime = createOverlayModalRuntimeService({
  getMainWindow: () => overlayManager.getMainWindow(),
  getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
});
const appState = createAppState({
  mpvSocketPath: getDefaultSocketPath(),
  texthookerPort: DEFAULT_TEXTHOOKER_PORT,
});
const immersionMediaRuntime = createImmersionMediaRuntime({
  getResolvedConfig: () => getResolvedConfig(),
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  getTracker: () => appState.immersionTracker,
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  logDebug: (message) => logger.debug(message),
  logInfo: (message) => logger.info(message),
});
const anilistStateRuntime = createAnilistStateRuntime({
  getClientSecretState: () => appState.anilistClientSecretState,
  setClientSecretState: (next) => {
    appState.anilistClientSecretState = next;
  },
  getRetryQueueState: () => appState.anilistRetryQueueState,
  setRetryQueueState: (next) => {
    appState.anilistRetryQueueState = next;
  },
  getUpdateQueueSnapshot: () => anilistUpdateQueue.getSnapshot(),
  clearStoredToken: () => anilistTokenStore.clearToken(),
  clearCachedAccessToken: () => {
    anilistCachedAccessToken = null;
  },
});
const configDerivedRuntime = createConfigDerivedRuntime({
  getResolvedConfig: () => getResolvedConfig(),
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
  platform: process.platform,
  defaultJimakuLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  defaultJimakuMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
  defaultJimakuApiBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
});
const subsyncRuntime = createMainSubsyncRuntime({
  getMpvClient: () => appState.mpvClient,
  getResolvedConfig: () => getResolvedConfig(),
  getSubsyncInProgress: () => appState.subsyncInProgress,
  setSubsyncInProgress: (inProgress) => {
    appState.subsyncInProgress = inProgress;
  },
  showMpvOsd: (text) => showMpvOsd(text),
  openManualPicker: (payload) => {
    sendToActiveOverlayWindow('subsync:open-manual', payload, {
      restoreOnModalClose: 'subsync',
    });
  },
});
let appTray: Tray | null = null;
const subtitleProcessingController = createSubtitleProcessingController({
  tokenizeSubtitle: async (text: string) => {
    if (getOverlayWindows().length === 0 && !subtitleWsService.hasClients()) {
      return null;
    }
    return await tokenizeSubtitle(text);
  },
  emitSubtitle: (payload) => {
    broadcastToOverlayWindows('subtitle:set', payload);
    subtitleWsService.broadcast(payload, {
      enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
      topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
      mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
    });
  },
  logDebug: (message) => {
    logger.debug(`[subtitle-processing] ${message}`);
  },
  now: () => Date.now(),
});
const overlayShortcutsRuntime = createOverlayShortcutsRuntimeService({
  getConfiguredShortcuts: () => getConfiguredShortcuts(),
  getShortcutsRegistered: () => appState.shortcutsRegistered,
  setShortcutsRegistered: (registered) => {
    appState.shortcutsRegistered = registered;
  },
  isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
  showMpvOsd: (text: string) => showMpvOsd(text),
  openRuntimeOptionsPalette: () => {
    openRuntimeOptionsPalette();
  },
  openJimaku: () => {
    sendToActiveOverlayWindow('jimaku:open', undefined, {
      restoreOnModalClose: 'jimaku',
    });
  },
  markAudioCard: () => markLastCardAsAudioCard(),
  copySubtitleMultiple: (timeoutMs) => {
    startPendingMultiCopy(timeoutMs);
  },
  copySubtitle: () => {
    copyCurrentSubtitle();
  },
  toggleSecondarySubMode: () => cycleSecondarySubMode(),
  updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
  triggerFieldGrouping: () => triggerFieldGrouping(),
  triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
  mineSentenceCard: () => mineSentenceCard(),
  mineSentenceMultiple: (timeoutMs) => {
    startPendingMineSentenceMultiple(timeoutMs);
  },
  cancelPendingMultiCopy: () => {
    cancelPendingMultiCopy();
  },
  cancelPendingMineSentenceMultiple: () => {
    cancelPendingMineSentenceMultiple();
  },
});

const notifyConfigHotReloadMessage = createConfigHotReloadMessageHandler({
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
});
const handleJellyfinRemotePlay = createHandleJellyfinRemotePlay({
  getConfiguredSession: () => getConfiguredJellyfinSession(getResolvedJellyfinConfig()),
  getClientInfo: () => getJellyfinClientInfo(),
  getJellyfinConfig: () => getResolvedJellyfinConfig(),
  playJellyfinItem: (params) =>
    playJellyfinItemInMpv(params as Parameters<typeof playJellyfinItemInMpv>[0]),
  logWarn: (message) => logger.warn(message),
});
const handleJellyfinRemotePlaystate = createHandleJellyfinRemotePlaystate({
  getMpvClient: () => appState.mpvClient,
  sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
  reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
  reportJellyfinRemoteStopped: () => reportJellyfinRemoteStopped(),
  jellyfinTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
});
const handleJellyfinRemoteGeneralCommand = createHandleJellyfinRemoteGeneralCommand({
  getMpvClient: () => appState.mpvClient,
  sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
  getActivePlayback: () => activeJellyfinRemotePlayback,
  reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
  logDebug: (message) => logger.debug(message),
});
const reportJellyfinRemoteProgress = createReportJellyfinRemoteProgressHandler({
  getActivePlayback: () => activeJellyfinRemotePlayback,
  clearActivePlayback: () => {
    activeJellyfinRemotePlayback = null;
  },
  getSession: () => appState.jellyfinRemoteSession,
  getMpvClient: () => appState.mpvClient,
  getNow: () => Date.now(),
  getLastProgressAtMs: () => jellyfinRemoteLastProgressAtMs,
  setLastProgressAtMs: (value) => {
    jellyfinRemoteLastProgressAtMs = value;
  },
  progressIntervalMs: JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS,
  ticksPerSecond: JELLYFIN_TICKS_PER_SECOND,
  logDebug: (message, error) => logger.debug(message, error),
});
const reportJellyfinRemoteStopped = createReportJellyfinRemoteStoppedHandler({
  getActivePlayback: () => activeJellyfinRemotePlayback,
  clearActivePlayback: () => {
    activeJellyfinRemotePlayback = null;
  },
  getSession: () => appState.jellyfinRemoteSession,
  logDebug: (message, error) => logger.debug(message, error),
});

const configHotReloadRuntime = createConfigHotReloadRuntime({
  getCurrentConfig: () => getResolvedConfig(),
  reloadConfigStrict: () => configService.reloadConfigStrict(),
  watchConfigPath: (configPath, onChange) => {
    const watchTarget = fs.existsSync(configPath) ? configPath : path.dirname(configPath);
    const watcher = fs.watch(watchTarget, (_eventType, filename) => {
      if (watchTarget === configPath) {
        onChange();
        return;
      }

      const normalized =
        typeof filename === 'string' ? filename : filename ? String(filename) : undefined;
      if (!normalized || normalized === 'config.json' || normalized === 'config.jsonc') {
        onChange();
      }
    });
    return {
      close: () => {
        watcher.close();
      },
    };
  },
  setTimeout: (callback, delayMs) => setTimeout(callback, delayMs),
  clearTimeout: (timeout) => clearTimeout(timeout),
  debounceMs: 250,
  onHotReloadApplied: createConfigHotReloadAppliedHandler({
    setKeybindings: (keybindings) => {
      appState.keybindings = keybindings;
    },
    refreshGlobalAndOverlayShortcuts: () => {
      refreshGlobalAndOverlayShortcuts();
    },
    setSecondarySubMode: (mode) => {
      appState.secondarySubMode = mode;
    },
    broadcastToOverlayWindows: (channel, payload) => {
      broadcastToOverlayWindows(channel, payload);
    },
    applyAnkiRuntimeConfigPatch: (patch) => {
      if (appState.ankiIntegration) {
        appState.ankiIntegration.applyRuntimeConfigPatch(patch);
      }
    },
  }),
  onRestartRequired: (fields) => notifyConfigHotReloadMessage(buildRestartRequiredConfigMessage(fields)),
  onInvalidConfig: notifyConfigHotReloadMessage,
  onValidationWarnings: (configPath, warnings) => {
    showDesktopNotification('SubMiner', {
      body: buildConfigWarningNotificationBody(configPath, warnings),
    });
  },
});

const jlptDictionaryRuntime = createJlptDictionaryRuntimeService({
  isJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
  getSearchPaths: () =>
    getJlptDictionarySearchPaths({
      getDictionaryRoots: () => [
        path.join(__dirname, '..', '..', 'vendor', 'yomitan-jlpt-vocab'),
        path.join(app.getAppPath(), 'vendor', 'yomitan-jlpt-vocab'),
        path.join(process.resourcesPath, 'yomitan-jlpt-vocab'),
        path.join(process.resourcesPath, 'app.asar', 'vendor', 'yomitan-jlpt-vocab'),
        USER_DATA_PATH,
        app.getPath('userData'),
        path.join(os.homedir(), '.config', 'SubMiner'),
        path.join(os.homedir(), '.config', 'subminer'),
        path.join(os.homedir(), 'Library', 'Application Support', 'SubMiner'),
        path.join(os.homedir(), 'Library', 'Application Support', 'subminer'),
        process.cwd(),
      ],
    }),
  setJlptLevelLookup: (lookup) => {
    appState.jlptLevelLookup = lookup;
  },
  log: (message) => {
    logger.info(`[JLPT] ${message}`);
  },
});

const frequencyDictionaryRuntime = createFrequencyDictionaryRuntimeService({
  isFrequencyDictionaryEnabled: () => getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
  getSearchPaths: () =>
    getFrequencyDictionarySearchPaths({
      getDictionaryRoots: () =>
        [
          path.join(__dirname, '..', '..', 'vendor', 'jiten_freq_global'),
          path.join(__dirname, '..', '..', 'vendor', 'frequency-dictionary'),
          path.join(app.getAppPath(), 'vendor', 'jiten_freq_global'),
          path.join(app.getAppPath(), 'vendor', 'frequency-dictionary'),
          path.join(process.resourcesPath, 'jiten_freq_global'),
          path.join(process.resourcesPath, 'frequency-dictionary'),
          path.join(process.resourcesPath, 'app.asar', 'vendor', 'jiten_freq_global'),
          path.join(process.resourcesPath, 'app.asar', 'vendor', 'frequency-dictionary'),
          USER_DATA_PATH,
          app.getPath('userData'),
          path.join(os.homedir(), '.config', 'SubMiner'),
          path.join(os.homedir(), '.config', 'subminer'),
          path.join(os.homedir(), 'Library', 'Application Support', 'SubMiner'),
          path.join(os.homedir(), 'Library', 'Application Support', 'subminer'),
          process.cwd(),
        ].filter((dictionaryRoot) => dictionaryRoot),
      getSourcePath: () => getResolvedConfig().subtitleStyle.frequencyDictionary.sourcePath,
    }),
  setFrequencyRankLookup: (lookup) => {
    appState.frequencyRankLookup = lookup;
  },
  log: (message) => {
    logger.info(`[Frequency] ${message}`);
  },
});

function getFieldGroupingResolver(): ((choice: KikuFieldGroupingChoice) => void) | null {
  return appState.fieldGroupingResolver;
}

function setFieldGroupingResolver(
  resolver: ((choice: KikuFieldGroupingChoice) => void) | null,
): void {
  if (!resolver) {
    appState.fieldGroupingResolver = null;
    return;
  }
  const sequence = ++appState.fieldGroupingResolverSequence;
  const wrappedResolver = (choice: KikuFieldGroupingChoice): void => {
    if (sequence !== appState.fieldGroupingResolverSequence) return;
    resolver(choice);
  };
  appState.fieldGroupingResolver = wrappedResolver;
}

const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntime<OverlayHostedModal>({
  getMainWindow: () => overlayManager.getMainWindow(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
  setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  setInvisibleOverlayVisible: (visible) => setInvisibleOverlayVisible(visible),
  getResolver: () => getFieldGroupingResolver(),
  setResolver: (resolver) => setFieldGroupingResolver(resolver),
  getRestoreVisibleOverlayOnModalClose: () =>
    overlayModalRuntime.getRestoreVisibleOverlayOnModalClose(),
  sendToVisibleOverlay: (channel, payload, runtimeOptions) => {
    return overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions);
  },
});
const createFieldGroupingCallback = fieldGroupingOverlayRuntime.createFieldGroupingCallback;

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
    broadcastToOverlayWindows('subtitle-position:set', position);
  },
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  setCurrentMediaTitle: (title) => {
    appState.currentMediaTitle = title;
  },
});

const overlayVisibilityRuntime = createOverlayVisibilityRuntimeService({
  getMainWindow: () => overlayManager.getMainWindow(),
  getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
  getWindowTracker: () => appState.windowTracker,
  getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
  setTrackerNotReadyWarningShown: (shown: boolean) => {
    appState.trackerNotReadyWarningShown = shown;
  },
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => updateVisibleOverlayBounds(geometry),
  updateInvisibleOverlayBounds: (geometry: WindowGeometry) =>
    updateInvisibleOverlayBounds(geometry),
  ensureOverlayWindowLevel: (window) => {
    ensureOverlayWindowLevel(window);
  },
  enforceOverlayLayerOrder: () => {
    enforceOverlayLayerOrder();
  },
  syncOverlayShortcuts: () => {
    overlayShortcutsRuntime.syncOverlayShortcuts();
  },
});

function getRuntimeOptionsState(): RuntimeOptionState[] {
  if (!appState.runtimeOptionsManager) return [];
  return appState.runtimeOptionsManager.listOptions();
}

function getOverlayWindows(): BrowserWindow[] {
  return overlayManager.getOverlayWindows();
}

function restorePreviousSecondarySubVisibility(): void {
  if (!appState.mpvClient || !appState.mpvClient.connected) return;
  appState.mpvClient.restorePreviousSecondarySubVisibility();
}

function broadcastToOverlayWindows(channel: string, ...args: unknown[]): void {
  overlayManager.broadcastToOverlayWindows(channel, ...args);
}

function broadcastRuntimeOptionsChanged(): void {
  broadcastRuntimeOptionsChangedRuntime(
    () => getRuntimeOptionsState(),
    (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  );
}

function sendToActiveOverlayWindow(
  channel: string,
  payload?: unknown,
  runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
): boolean {
  return overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions);
}

function setOverlayDebugVisualizationEnabled(enabled: boolean): void {
  setOverlayDebugVisualizationEnabledRuntime(
    appState.overlayDebugVisualizationEnabled,
    enabled,
    (next) => {
      appState.overlayDebugVisualizationEnabled = next;
    },
    (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  );
}

function openRuntimeOptionsPalette(): void {
  overlayModalRuntime.openRuntimeOptionsPalette();
}

function getResolvedConfig() {
  return configService.getConfig();
}

function getResolvedJellyfinConfig() {
  return getResolvedConfig().jellyfin;
}

function getJellyfinClientInfo(config = getResolvedJellyfinConfig()) {
  const clientName = config.clientName || DEFAULT_CONFIG.jellyfin.clientName;
  const clientVersion = config.clientVersion || DEFAULT_CONFIG.jellyfin.clientVersion;
  const deviceId = config.deviceId || DEFAULT_CONFIG.jellyfin.deviceId;
  return {
    clientName,
    clientVersion,
    deviceId,
  };
}

const waitForMpvConnected = createWaitForMpvConnectedHandler({
  getMpvClient: () => appState.mpvClient,
  now: () => Date.now(),
  sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
});

const launchMpvIdleForJellyfinPlayback = createLaunchMpvIdleForJellyfinPlaybackHandler({
  getSocketPath: () => appState.mpvSocketPath,
  platform: process.platform,
  execPath: process.execPath,
  defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
  defaultMpvArgs: MPV_JELLYFIN_DEFAULT_ARGS,
  removeSocketPath: (socketPath) => {
    fs.rmSync(socketPath, { force: true });
  },
  spawnMpv: (args) =>
    spawn('mpv', args, {
      detached: true,
      stdio: 'ignore',
    }),
  logWarn: (message, error) => logger.warn(message, error),
  logInfo: (message) => logger.info(message),
});

const ensureMpvConnectedForJellyfinPlayback = createEnsureMpvConnectedForJellyfinPlaybackHandler({
  getMpvClient: () => appState.mpvClient,
  setMpvClient: (client) => {
    appState.mpvClient = client as MpvIpcClient | null;
  },
  createMpvClient: () => createMpvClientRuntimeService(),
  waitForMpvConnected: (timeoutMs) => waitForMpvConnected(timeoutMs),
  launchMpvIdleForJellyfinPlayback: () => launchMpvIdleForJellyfinPlayback(),
  getAutoLaunchInFlight: () => jellyfinMpvAutoLaunchInFlight,
  setAutoLaunchInFlight: (promise) => {
    jellyfinMpvAutoLaunchInFlight = promise;
  },
  connectTimeoutMs: JELLYFIN_MPV_CONNECT_TIMEOUT_MS,
  autoLaunchTimeoutMs: JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS,
});

async function playJellyfinItemInMpv(params: {
  session: {
    serverUrl: string;
    accessToken: string;
    userId: string;
    username: string;
  };
  clientInfo: ReturnType<typeof getJellyfinClientInfo>;
  jellyfinConfig: ReturnType<typeof getResolvedJellyfinConfig>;
  itemId: string;
  audioStreamIndex?: number;
  subtitleStreamIndex?: number;
  startTimeTicksOverride?: number;
  setQuitOnDisconnectArm?: boolean;
}): Promise<void> {
  const connected = await ensureMpvConnectedForJellyfinPlayback();
  if (!connected || !appState.mpvClient) {
    throw new Error(
      'MPV not connected and auto-launch failed. Ensure mpv is installed and available in PATH.',
    );
  }

  const plan = await resolveJellyfinPlaybackPlanRuntime(
    params.session,
    params.clientInfo,
    params.jellyfinConfig,
    {
      itemId: params.itemId,
      audioStreamIndex: params.audioStreamIndex,
      subtitleStreamIndex: params.subtitleStreamIndex,
    },
  );

  applyJellyfinMpvDefaults(appState.mpvClient);
  sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'sub-auto', 'no']);
  sendMpvCommandRuntime(appState.mpvClient, ['loadfile', plan.url, 'replace']);
  if (params.setQuitOnDisconnectArm !== false) {
    jellyfinPlayQuitOnDisconnectArmed = false;
    setTimeout(() => {
      jellyfinPlayQuitOnDisconnectArmed = true;
    }, 3000);
  }
  sendMpvCommandRuntime(appState.mpvClient, [
    'set_property',
    'force-media-title',
    `[Jellyfin/${plan.mode}] ${plan.title}`,
  ]);
  sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'sid', 'no']);
  setTimeout(() => {
    sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'sid', 'no']);
  }, 500);

  const startTimeTicks =
    typeof params.startTimeTicksOverride === 'number'
      ? Math.max(0, params.startTimeTicksOverride)
      : plan.startTimeTicks;
  if (startTimeTicks > 0) {
    sendMpvCommandRuntime(appState.mpvClient, [
      'seek',
      jellyfinTicksToSecondsRuntime(startTimeTicks),
      'absolute+exact',
    ]);
  }

  void (async () => {
    try {
      const normalizeLang = (value: unknown): string =>
        String(value || '')
          .trim()
          .toLowerCase()
          .replace(/_/g, '-');
      const isJapanese = (value: string): boolean => {
        const v = normalizeLang(value);
        return (
          v === 'ja' ||
          v === 'jp' ||
          v === 'jpn' ||
          v === 'japanese' ||
          v.startsWith('ja-') ||
          v.startsWith('jp-')
        );
      };
      const isEnglish = (value: string): boolean => {
        const v = normalizeLang(value);
        return (
          v === 'en' ||
          v === 'eng' ||
          v === 'english' ||
          v === 'enus' ||
          v === 'en-us' ||
          v.startsWith('en-')
        );
      };
      const isLikelyHearingImpaired = (title: string): boolean =>
        /\b(hearing impaired|sdh|closed captions?|cc)\b/i.test(title);
      const pickBestTrackId = (
        tracks: Array<{
          id: number;
          lang: string;
          title: string;
          external: boolean;
        }>,
        languageMatcher: (value: string) => boolean,
        excludeId: number | null = null,
      ): number | null => {
        const ranked = tracks
          .filter((track) => languageMatcher(track.lang))
          .filter((track) => track.id !== excludeId)
          .map((track) => ({
            track,
            score:
              (track.external ? 100 : 0) +
              (isLikelyHearingImpaired(track.title) ? -10 : 10) +
              (/\bdefault\b/i.test(track.title) ? 3 : 0),
          }))
          .sort((a, b) => b.score - a.score);
        return ranked[0]?.track.id ?? null;
      };

      const tracks = await listJellyfinSubtitleTracksRuntime(
        params.session,
        params.clientInfo,
        params.itemId,
      );
      const externalTracks = tracks.filter((track) => Boolean(track.deliveryUrl));

      if (externalTracks.length === 0) return;
      await new Promise<void>((resolve) => setTimeout(resolve, 300));
      const seenUrls = new Set<string>();
      for (const track of externalTracks) {
        if (!track.deliveryUrl) continue;
        if (seenUrls.has(track.deliveryUrl)) continue;
        seenUrls.add(track.deliveryUrl);
        const labelBase = (track.title || track.language || '').trim();
        const label = labelBase || `Jellyfin Subtitle ${track.index}`;
        sendMpvCommandRuntime(appState.mpvClient, [
          'sub-add',
          track.deliveryUrl,
          'cached',
          label,
          track.language || '',
        ]);
      }
      await new Promise<void>((resolve) => setTimeout(resolve, 250));
      const trackListRaw = await appState.mpvClient?.requestProperty('track-list');
      const subtitleTracks = Array.isArray(trackListRaw)
        ? trackListRaw
            .filter(
              (track): track is Record<string, unknown> =>
                Boolean(track) &&
                typeof track === 'object' &&
                track.type === 'sub' &&
                typeof track.id === 'number',
            )
            .map((track) => ({
              id: track.id as number,
              lang: String(track.lang || ''),
              title: String(track.title || ''),
              external: track.external === true,
            }))
        : [];

      const japanesePrimaryId = pickBestTrackId(subtitleTracks, isJapanese);
      if (japanesePrimaryId !== null) {
        sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'sid', japanesePrimaryId]);
      } else {
        sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'sid', 'no']);
      }

      const englishSecondaryId = pickBestTrackId(subtitleTracks, isEnglish, japanesePrimaryId);
      if (englishSecondaryId !== null) {
        sendMpvCommandRuntime(appState.mpvClient, [
          'set_property',
          'secondary-sid',
          englishSecondaryId,
        ]);
      }
    } catch (error) {
      logger.debug('Failed to preload Jellyfin external subtitles', error);
    }
  })();

  activeJellyfinRemotePlayback = {
    itemId: params.itemId,
    mediaSourceId: undefined,
    audioStreamIndex: plan.audioStreamIndex,
    subtitleStreamIndex: plan.subtitleStreamIndex,
    playMethod: plan.mode === 'direct' ? 'DirectPlay' : 'Transcode',
  };
  jellyfinRemoteLastProgressAtMs = 0;
  void appState.jellyfinRemoteSession?.reportPlaying({
    itemId: params.itemId,
    mediaSourceId: undefined,
    playMethod: activeJellyfinRemotePlayback.playMethod,
    audioStreamIndex: plan.audioStreamIndex,
    subtitleStreamIndex: plan.subtitleStreamIndex,
    eventName: 'start',
  });
  showMpvOsd(`Jellyfin ${plan.mode}: ${plan.title}`);
}

const handleJellyfinAuthCommands = createHandleJellyfinAuthCommands({
  patchRawConfig: (patch) => {
    configService.patchRawConfig(patch);
  },
  authenticateWithPassword: (serverUrl, username, password, clientInfo) =>
    authenticateWithPasswordRuntime(serverUrl, username, password, clientInfo),
  logInfo: (message) => logger.info(message),
});

const handleJellyfinListCommands = createHandleJellyfinListCommands({
  listJellyfinLibraries: (session, clientInfo) => listJellyfinLibrariesRuntime(session, clientInfo),
  listJellyfinItems: (session, clientInfo, params) =>
    listJellyfinItemsRuntime(session, clientInfo, params),
  listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
    listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
  logInfo: (message) => logger.info(message),
});

const handleJellyfinPlayCommand = createHandleJellyfinPlayCommand({
  playJellyfinItemInMpv: (params) =>
    playJellyfinItemInMpv(params as Parameters<typeof playJellyfinItemInMpv>[0]),
  logWarn: (message) => logger.warn(message),
});

const handleJellyfinRemoteAnnounceCommand = createHandleJellyfinRemoteAnnounceCommand({
  startJellyfinRemoteSession: () => startJellyfinRemoteSession(),
  getRemoteSession: () => appState.jellyfinRemoteSession,
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
});

const startJellyfinRemoteSession = createStartJellyfinRemoteSessionHandler({
  getJellyfinConfig: () => getResolvedJellyfinConfig(),
  getCurrentSession: () => appState.jellyfinRemoteSession,
  setCurrentSession: (session) => {
    appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
  },
  createRemoteSessionService: (options) => new JellyfinRemoteSessionService(options),
  defaultDeviceId: DEFAULT_CONFIG.jellyfin.deviceId,
  defaultClientName: DEFAULT_CONFIG.jellyfin.clientName,
  defaultClientVersion: DEFAULT_CONFIG.jellyfin.clientVersion,
  handlePlay: (payload) => handleJellyfinRemotePlay(payload),
  handlePlaystate: (payload) => handleJellyfinRemotePlaystate(payload),
  handleGeneralCommand: (payload) => handleJellyfinRemoteGeneralCommand(payload),
  logInfo: (message) => logger.info(message),
  logWarn: (message, details) => logger.warn(message, details),
});

const stopJellyfinRemoteSession = createStopJellyfinRemoteSessionHandler({
  getCurrentSession: () => appState.jellyfinRemoteSession,
  setCurrentSession: (session) => {
    appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
  },
  clearActivePlayback: () => {
    activeJellyfinRemotePlayback = null;
  },
});

async function runJellyfinCommand(args: CliArgs): Promise<void> {
  const jellyfinConfig = getResolvedJellyfinConfig();
  const serverUrl =
    args.jellyfinServer?.trim() || jellyfinConfig.serverUrl || DEFAULT_CONFIG.jellyfin.serverUrl;
  const clientInfo = getJellyfinClientInfo(jellyfinConfig);

  if (
    await handleJellyfinAuthCommands({
      args,
      jellyfinConfig,
      serverUrl,
      clientInfo,
    })
  ) {
    return;
  }

  const accessToken = jellyfinConfig.accessToken;
  const userId = jellyfinConfig.userId;
  if (!serverUrl || !accessToken || !userId) {
    throw new Error('Missing Jellyfin session. Run --jellyfin-login first.');
  }
  const session = {
    serverUrl,
    accessToken,
    userId,
    username: jellyfinConfig.username,
  };

  if (await handleJellyfinRemoteAnnounceCommand(args)) {
    return;
  }

  if (
    await handleJellyfinListCommands({
      args,
      session,
      clientInfo,
      jellyfinConfig,
    })
  ) {
    return;
  }

  if (
    await handleJellyfinPlayCommand({
      args,
      session,
      clientInfo,
      jellyfinConfig,
    })
  ) {
    return;
  }
}

const notifyAnilistSetup = createNotifyAnilistSetupHandler({
  hasMpvClient: () => Boolean(appState.mpvClient),
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  logInfo: (message) => logger.info(message),
});

const consumeAnilistSetupTokenFromUrl = createConsumeAnilistSetupTokenFromUrlHandler({
  consumeAnilistSetupCallbackUrl,
  saveToken: (token) => anilistTokenStore.saveToken(token),
  setCachedToken: (token) => {
    anilistCachedAccessToken = token;
  },
  setResolvedState: (resolvedAt) => {
    anilistStateRuntime.setClientSecretState({
      status: 'resolved',
      source: 'stored',
      message: 'saved token from AniList login',
      resolvedAt,
      errorAt: null,
    });
  },
  setSetupPageOpened: (opened) => {
    appState.anilistSetupPageOpened = opened;
  },
  onSuccess: () => {
    notifyAnilistSetup('AniList login success');
  },
  closeWindow: () => {
    if (appState.anilistSetupWindow && !appState.anilistSetupWindow.isDestroyed()) {
      appState.anilistSetupWindow.close();
    }
  },
});

const handleAnilistSetupProtocolUrl = createHandleAnilistSetupProtocolUrlHandler({
  consumeAnilistSetupTokenFromUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
  logWarn: (message, details) => logger.warn(message, details),
});

const registerSubminerProtocolClient = createRegisterSubminerProtocolClientHandler({
  isDefaultApp: () => Boolean(process.defaultApp),
  getArgv: () => process.argv,
  execPath: process.execPath,
  resolvePath: (value) => path.resolve(value),
  setAsDefaultProtocolClient: (scheme, appPath, args) =>
    appPath ? app.setAsDefaultProtocolClient(scheme, appPath, args) : app.setAsDefaultProtocolClient(scheme),
  logWarn: (message, details) => logger.warn(message, details),
});

function openAnilistSetupWindow(): void {
  const maybeFocusExistingAnilistSetupWindow = createMaybeFocusExistingAnilistSetupWindowHandler({
    getSetupWindow: () => appState.anilistSetupWindow,
  });
  if (maybeFocusExistingAnilistSetupWindow()) {
    return;
  }

  const setupWindow = new BrowserWindow({
    width: 1000,
    height: 760,
    title: 'Anilist Setup',
    show: true,
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });
  const authorizeUrl = buildAnilistSetupUrl({
    authorizeUrl: ANILIST_SETUP_CLIENT_ID_URL,
    clientId: ANILIST_DEFAULT_CLIENT_ID,
    responseType: ANILIST_SETUP_RESPONSE_TYPE,
  });
  const consumeCallbackUrl = (rawUrl: string): boolean => consumeAnilistSetupTokenFromUrl(rawUrl);
  const openSetupInBrowser = () =>
    openAnilistSetupInBrowser({
      authorizeUrl,
      openExternal: (url) => shell.openExternal(url),
      logError: (message, error) => logger.error(message, error),
    });
  const loadManualTokenEntry = () =>
    loadAnilistManualTokenEntry({
      setupWindow,
      authorizeUrl,
      developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
      logWarn: (message, data) => logger.warn(message, data),
    });
  const handleManualAnilistSetupSubmission = createHandleManualAnilistSetupSubmissionHandler({
    consumeCallbackUrl: (rawUrl) => consumeCallbackUrl(rawUrl),
    redirectUri: ANILIST_REDIRECT_URI,
    logWarn: (message) => logger.warn(message),
  });
  const anilistSetupFallback = createAnilistSetupFallbackHandler({
    authorizeUrl,
    developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
    setupWindow,
    openSetupInBrowser,
    loadManualTokenEntry,
    logError: (message, details) => logger.error(message, details),
    logWarn: (message) => logger.warn(message),
  });
  const handleAnilistSetupWindowOpen = createAnilistSetupWindowOpenHandler({
    isAllowedExternalUrl: (url) => isAllowedAnilistExternalUrl(url),
    openExternal: (url) => {
      void shell.openExternal(url);
    },
    logWarn: (message, details) => logger.warn(message, details),
  });
  const handleAnilistSetupWillNavigate = createAnilistSetupWillNavigateHandler({
    handleManualSubmission: (url) => handleManualAnilistSetupSubmission(url),
    consumeCallbackUrl: (url) => consumeCallbackUrl(url),
    redirectUri: ANILIST_REDIRECT_URI,
    isAllowedNavigationUrl: (url) => isAllowedAnilistSetupNavigationUrl(url),
    logWarn: (message, details) => logger.warn(message, details),
  });
  const handleAnilistSetupWillRedirect = createAnilistSetupWillRedirectHandler({
    consumeCallbackUrl: (url) => consumeCallbackUrl(url),
  });
  const handleAnilistSetupDidNavigate = createAnilistSetupDidNavigateHandler({
    consumeCallbackUrl: (url) => consumeCallbackUrl(url),
  });
  const handleAnilistSetupDidFailLoad = createAnilistSetupDidFailLoadHandler({
    onLoadFailure: (details) => anilistSetupFallback.onLoadFailure(details),
  });
  const handleAnilistSetupDidFinishLoad = createAnilistSetupDidFinishLoadHandler({
    getLoadedUrl: () => setupWindow.webContents.getURL(),
    onBlankPageLoaded: () => anilistSetupFallback.onBlankPageLoaded(),
  });
  const handleAnilistSetupWindowClosed = createHandleAnilistSetupWindowClosedHandler({
    clearSetupWindow: () => {
      appState.anilistSetupWindow = null;
    },
    setSetupPageOpened: (opened) => {
      appState.anilistSetupPageOpened = opened;
    },
  });
  const handleAnilistSetupWindowOpened = createHandleAnilistSetupWindowOpenedHandler({
    setSetupWindow: () => {
      appState.anilistSetupWindow = setupWindow;
    },
    setSetupPageOpened: (opened) => {
      appState.anilistSetupPageOpened = opened;
    },
  });

  setupWindow.webContents.setWindowOpenHandler(({ url }) => handleAnilistSetupWindowOpen({ url }));
  setupWindow.webContents.on('will-navigate', (event, url) => {
    handleAnilistSetupWillNavigate({
      url,
      preventDefault: () => event.preventDefault(),
    });
  });
  setupWindow.webContents.on('will-redirect', (event, url) => {
    handleAnilistSetupWillRedirect({
      url,
      preventDefault: () => event.preventDefault(),
    });
  });
  setupWindow.webContents.on('did-navigate', (_event, url) => {
    handleAnilistSetupDidNavigate(url);
  });

  setupWindow.webContents.on(
    'did-fail-load',
    (_event, errorCode, errorDescription, validatedURL) => {
      handleAnilistSetupDidFailLoad({
        errorCode,
        errorDescription,
        validatedURL,
      });
    },
  );

  setupWindow.webContents.on('did-finish-load', () => {
    handleAnilistSetupDidFinishLoad();
  });

  loadManualTokenEntry();

  setupWindow.on('closed', () => {
    handleAnilistSetupWindowClosed();
  });

  handleAnilistSetupWindowOpened();
}

function openJellyfinSetupWindow(): void {
  const maybeFocusExistingJellyfinSetupWindow = createMaybeFocusExistingJellyfinSetupWindowHandler({
    getSetupWindow: () => appState.jellyfinSetupWindow,
  });
  if (maybeFocusExistingJellyfinSetupWindow()) {
    return;
  }

  const setupWindow = new BrowserWindow({
    width: 520,
    height: 560,
    title: 'Jellyfin Setup',
    show: true,
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  const defaults = getResolvedJellyfinConfig();
  const defaultServer = defaults.serverUrl || 'http://127.0.0.1:8096';
  const defaultUser = defaults.username || '';
  const formHtml = buildJellyfinSetupFormHtml(defaultServer, defaultUser);
  const handleJellyfinSetupSubmission = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: (server, username, password, clientInfo) =>
      authenticateWithPasswordRuntime(server, username, password, clientInfo),
    getJellyfinClientInfo: () => getJellyfinClientInfo(),
    patchJellyfinConfig: (session) => {
      configService.patchRawConfig({
        jellyfin: {
          enabled: true,
          serverUrl: session.serverUrl,
          username: session.username,
          accessToken: session.accessToken,
          userId: session.userId,
        },
      });
    },
    logInfo: (message) => logger.info(message),
    logError: (message, error) => logger.error(message, error),
    showMpvOsd: (message) => showMpvOsd(message),
    closeSetupWindow: () => {
      if (!setupWindow.isDestroyed()) {
        setupWindow.close();
      }
    },
  });
  const handleJellyfinSetupNavigation = createHandleJellyfinSetupNavigationHandler({
    setupSchemePrefix: 'subminer://jellyfin-setup',
    handleSubmission: (rawUrl) => handleJellyfinSetupSubmission(rawUrl),
    logError: (message, error) => logger.error(message, error),
  });
  const handleJellyfinSetupWindowClosed = createHandleJellyfinSetupWindowClosedHandler({
    clearSetupWindow: () => {
      appState.jellyfinSetupWindow = null;
    },
  });
  const handleJellyfinSetupWindowOpened = createHandleJellyfinSetupWindowOpenedHandler({
    setSetupWindow: () => {
      appState.jellyfinSetupWindow = setupWindow;
    },
  });

  setupWindow.webContents.on('will-navigate', (event, url) => {
    handleJellyfinSetupNavigation({
      url,
      preventDefault: () => event.preventDefault(),
    });
  });

  void setupWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(formHtml)}`);

  setupWindow.on('closed', () => {
    handleJellyfinSetupWindowClosed();
  });

  handleJellyfinSetupWindowOpened();
}

const refreshAnilistClientSecretState = createRefreshAnilistClientSecretStateHandler({
  getResolvedConfig: () => getResolvedConfig(),
  isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config),
  getCachedAccessToken: () => anilistCachedAccessToken,
  setCachedAccessToken: (token) => {
    anilistCachedAccessToken = token;
  },
  saveStoredToken: (token) => {
    anilistTokenStore.saveToken(token);
  },
  loadStoredToken: () => anilistTokenStore.loadToken(),
  setClientSecretState: (state) => {
    anilistStateRuntime.setClientSecretState(state);
  },
  getAnilistSetupPageOpened: () => appState.anilistSetupPageOpened,
  setAnilistSetupPageOpened: (opened) => {
    appState.anilistSetupPageOpened = opened;
  },
  openAnilistSetupWindow: () => {
    openAnilistSetupWindow();
  },
  now: () => Date.now(),
});

function getCurrentAnilistMediaKey(): string | null {
  const path = appState.currentMediaPath?.trim();
  return path && path.length > 0 ? path : null;
}

function resetAnilistMediaTracking(mediaKey: string | null): void {
  anilistCurrentMediaKey = mediaKey;
  anilistCurrentMediaDurationSec = null;
  anilistCurrentMediaGuess = null;
  anilistCurrentMediaGuessPromise = null;
  anilistLastDurationProbeAtMs = 0;
}

const getAnilistMediaGuessRuntimeState = () => ({
  mediaKey: anilistCurrentMediaKey,
  mediaDurationSec: anilistCurrentMediaDurationSec,
  mediaGuess: anilistCurrentMediaGuess,
  mediaGuessPromise: anilistCurrentMediaGuessPromise,
  lastDurationProbeAtMs: anilistLastDurationProbeAtMs,
});

const setAnilistMediaGuessRuntimeState = (state: {
  mediaKey: string | null;
  mediaDurationSec: number | null;
  mediaGuess: AnilistMediaGuess | null;
  mediaGuessPromise: Promise<AnilistMediaGuess | null> | null;
  lastDurationProbeAtMs: number;
}) => {
  anilistCurrentMediaKey = state.mediaKey;
  anilistCurrentMediaDurationSec = state.mediaDurationSec;
  anilistCurrentMediaGuess = state.mediaGuess;
  anilistCurrentMediaGuessPromise = state.mediaGuessPromise;
  anilistLastDurationProbeAtMs = state.lastDurationProbeAtMs;
};

const maybeProbeAnilistDuration = createMaybeProbeAnilistDurationHandler({
  getState: () => getAnilistMediaGuessRuntimeState(),
  setState: (state) => {
    setAnilistMediaGuessRuntimeState(state);
  },
  durationRetryIntervalMs: ANILIST_DURATION_RETRY_INTERVAL_MS,
  now: () => Date.now(),
  requestMpvDuration: async () => appState.mpvClient?.requestProperty('duration'),
  logWarn: (message, error) => logger.warn(message, error),
});

const ensureAnilistMediaGuess = createEnsureAnilistMediaGuessHandler({
  getState: () => getAnilistMediaGuessRuntimeState(),
  setState: (state) => {
    setAnilistMediaGuessRuntimeState(state);
  },
  resolveMediaPathForJimaku: (currentMediaPath) => mediaRuntime.resolveMediaPathForJimaku(currentMediaPath),
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
});

const rememberAnilistAttemptedUpdate = (key: string): void => {
  rememberAnilistAttemptedUpdateKey(anilistAttemptedUpdateKeys, key, ANILIST_MAX_ATTEMPTED_UPDATE_KEYS);
};

const processNextAnilistRetryUpdate = createProcessNextAnilistRetryUpdateHandler({
  nextReady: () => anilistUpdateQueue.nextReady(),
  refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
  setLastAttemptAt: (value) => {
    appState.anilistRetryQueueState.lastAttemptAt = value;
  },
  setLastError: (value) => {
    appState.anilistRetryQueueState.lastError = value;
  },
  refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
  updateAnilistPostWatchProgress: (accessToken, title, episode) =>
    updateAnilistPostWatchProgress(accessToken, title, episode),
  markSuccess: (key) => {
    anilistUpdateQueue.markSuccess(key);
  },
  rememberAttemptedUpdateKey: (key) => {
    rememberAnilistAttemptedUpdate(key);
  },
  markFailure: (key, message) => {
    anilistUpdateQueue.markFailure(key, message);
  },
  logInfo: (message) => logger.info(message),
  now: () => Date.now(),
});

const maybeRunAnilistPostWatchUpdate = createMaybeRunAnilistPostWatchUpdateHandler({
  getInFlight: () => anilistUpdateInFlight,
  setInFlight: (value) => {
    anilistUpdateInFlight = value;
  },
  getResolvedConfig: () => getResolvedConfig(),
  isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
  getCurrentMediaKey: () => getCurrentAnilistMediaKey(),
  hasMpvClient: () => Boolean(appState.mpvClient),
  getTrackedMediaKey: () => anilistCurrentMediaKey,
  resetTrackedMedia: (mediaKey) => {
    resetAnilistMediaTracking(mediaKey);
  },
  getWatchedSeconds: () => appState.mpvClient?.currentTimePos ?? Number.NaN,
  maybeProbeAnilistDuration: (mediaKey) => maybeProbeAnilistDuration(mediaKey),
  ensureAnilistMediaGuess: (mediaKey) => ensureAnilistMediaGuess(mediaKey),
  hasAttemptedUpdateKey: (key) => anilistAttemptedUpdateKeys.has(key),
  processNextAnilistRetryUpdate: () => processNextAnilistRetryUpdate(),
  refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
  enqueueRetry: (key, title, episode) => {
    anilistUpdateQueue.enqueue(key, title, episode);
  },
  markRetryFailure: (key, message) => {
    anilistUpdateQueue.markFailure(key, message);
  },
  markRetrySuccess: (key) => {
    anilistUpdateQueue.markSuccess(key);
  },
  refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
  updateAnilistPostWatchProgress: (accessToken, title, episode) =>
    updateAnilistPostWatchProgress(accessToken, title, episode),
  rememberAttemptedUpdateKey: (key) => {
    rememberAnilistAttemptedUpdate(key);
  },
  showMpvOsd: (message) => showMpvOsd(message),
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
  minWatchSeconds: ANILIST_UPDATE_MIN_WATCH_SECONDS,
  minWatchRatio: ANILIST_UPDATE_MIN_WATCH_RATIO,
});

const loadSubtitlePosition = createLoadSubtitlePositionHandler({
  loadSubtitlePositionCore: () =>
    loadSubtitlePositionCore({
      currentMediaPath: appState.currentMediaPath,
      fallbackPosition: getResolvedConfig().subtitlePosition,
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

registerSubminerProtocolClient();

registerProtocolUrlHandlers({
  registerOpenUrl: (listener) => {
    app.on('open-url', listener);
  },
  registerSecondInstance: (listener) => {
    app.on('second-instance', listener);
  },
  handleAnilistSetupProtocolUrl: (rawUrl) => handleAnilistSetupProtocolUrl(rawUrl),
  findAnilistSetupDeepLinkArgvUrl: (argv) => findAnilistSetupDeepLinkArgvUrl(argv),
  logUnhandledOpenUrl: (rawUrl) => {
    logger.warn('Unhandled app protocol URL', { rawUrl });
  },
  logUnhandledSecondInstanceUrl: (rawUrl) => {
    logger.warn('Unhandled second-instance protocol URL', { rawUrl });
  },
});

const startupState = runStartupBootstrapRuntime(
  createStartupBootstrapRuntimeDeps({
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
    getDefaultSocketPath: () => getDefaultSocketPath(),
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
    onConfigGenerated: (exitCode: number) => {
      process.exitCode = exitCode;
      app.quit();
    },
    onGenerateConfigError: (error: Error) => {
      logger.error(`Failed to generate config: ${error.message}`);
      process.exitCode = 1;
      app.quit();
    },
    startAppLifecycle: createAppLifecycleRuntimeRunner({
      app,
      platform: process.platform,
      shouldStartApp: (nextArgs: CliArgs) => shouldStartApp(nextArgs),
      parseArgs: (argv: string[]) => parseArgs(argv),
      handleCliCommand: (nextArgs: CliArgs, source: CliCommandSource) =>
        handleCliCommand(nextArgs, source),
      printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
      logNoRunningInstance: () => appLogger.logNoRunningInstance(),
      onReady: createAppReadyRuntimeRunner({
        loadSubtitlePosition: () => loadSubtitlePosition(),
        resolveKeybindings: () => {
          appState.keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);
        },
        createMpvClient: () => {
          appState.mpvClient = createMpvClientRuntimeService();
        },
        reloadConfig: createReloadConfigHandler({
          reloadConfigStrict: () => configService.reloadConfigStrict(),
          logInfo: (message) => appLogger.logInfo(message),
          logWarning: (message) => appLogger.logWarning(message),
          showDesktopNotification: (title, options) => showDesktopNotification(title, options),
          startConfigHotReload: () => configHotReloadRuntime.start(),
          refreshAnilistClientSecretState: (options) => refreshAnilistClientSecretState(options),
          failHandlers: {
            logError: (details) => logger.error(details),
            showErrorBox: (title, details) => dialog.showErrorBox(title, details),
            quit: () => app.quit(),
          },
        }),
        getResolvedConfig: () => getResolvedConfig(),
        getConfigWarnings: () => configService.getWarnings(),
        logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
        setLogLevel: (level: string, source: LogLevelSource) => setLogLevel(level, source),
        initRuntimeOptionsManager: () => {
          appState.runtimeOptionsManager = new RuntimeOptionsManager(
            () => configService.getConfig().ankiConnect,
            {
              applyAnkiPatch: (patch) => {
                if (appState.ankiIntegration) {
                  appState.ankiIntegration.applyRuntimeConfigPatch(patch);
                }
              },
              onOptionsChanged: () => {
                broadcastRuntimeOptionsChanged();
                refreshOverlayShortcuts();
              },
            },
          );
        },
        setSecondarySubMode: (mode: SecondarySubMode) => {
          appState.secondarySubMode = mode;
        },
        defaultSecondarySubMode: 'hover',
        defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
        hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
        startSubtitleWebsocket: (port: number) => {
          subtitleWsService.start(port, () => appState.currentSubText);
        },
        log: (message) => appLogger.logInfo(message),
        createMecabTokenizerAndCheck: async () => {
          await createMecabTokenizerAndCheck();
        },
        createSubtitleTimingTracker: () => {
          const tracker = new SubtitleTimingTracker();
          appState.subtitleTimingTracker = tracker;
        },
        createImmersionTracker: createImmersionTrackerStartupHandler({
          getResolvedConfig: () => getResolvedConfig(),
          getConfiguredDbPath: () => immersionMediaRuntime.getConfiguredDbPath(),
          createTrackerService: (params) => new ImmersionTrackerService(params),
          setTracker: (tracker) => {
            appState.immersionTracker = tracker as ImmersionTrackerService | null;
          },
          getMpvClient: () => appState.mpvClient,
          seedTrackerFromCurrentMedia: () => {
            void immersionMediaRuntime.seedFromCurrentMedia();
          },
          logInfo: (message) => logger.info(message),
          logDebug: (message) => logger.debug(message),
          logWarn: (message, details) => logger.warn(message, details),
        }),
        loadYomitanExtension: async () => {
          await loadYomitanExtension();
        },
        startJellyfinRemoteSession: async () => {
          await startJellyfinRemoteSession();
        },
        prewarmSubtitleDictionaries: async () => {
          await prewarmSubtitleDictionaries();
        },
        startBackgroundWarmups: () => {
          startBackgroundWarmups();
        },
        texthookerOnlyMode: appState.texthookerOnlyMode,
        shouldAutoInitializeOverlayRuntimeFromConfig: () =>
          appState.backgroundMode
            ? false
            : configDerivedRuntime.shouldAutoInitializeOverlayRuntimeFromConfig(),
        initializeOverlayRuntime: () => initializeOverlayRuntime(),
        handleInitialArgs: () => handleInitialArgs(),
        onCriticalConfigErrors: createCriticalConfigErrorHandler({
          getConfigPath: () => configService.getConfigPath(),
          failHandlers: {
            logError: (message) => logger.error(message),
            showErrorBox: (title, message) => dialog.showErrorBox(title, message),
            quit: () => app.quit(),
          },
        }),
        logDebug: (message: string) => {
          logger.debug(message);
        },
        now: () => Date.now(),
      }),
      onWillQuitCleanup: () => {
        destroyTray();
        configHotReloadRuntime.stop();
        restorePreviousSecondarySubVisibility();
        globalShortcut.unregisterAll();
        subtitleWsService.stop();
        texthookerService.stop();
        if (appState.yomitanParserWindow && !appState.yomitanParserWindow.isDestroyed()) {
          appState.yomitanParserWindow.destroy();
        }
        appState.yomitanParserWindow = null;
        appState.yomitanParserReadyPromise = null;
        appState.yomitanParserInitPromise = null;
        if (appState.windowTracker) {
          appState.windowTracker.stop();
        }
        if (appState.mpvClient && appState.mpvClient.socket) {
          appState.mpvClient.socket.destroy();
        }
        if (appState.reconnectTimer) {
          clearTimeout(appState.reconnectTimer);
        }
        if (appState.subtitleTimingTracker) {
          appState.subtitleTimingTracker.destroy();
        }
        if (appState.immersionTracker) {
          appState.immersionTracker.destroy();
          appState.immersionTracker = null;
        }
        if (appState.ankiIntegration) {
          appState.ankiIntegration.destroy();
        }
        if (appState.anilistSetupWindow) {
          appState.anilistSetupWindow.destroy();
        }
        appState.anilistSetupWindow = null;
        if (appState.jellyfinSetupWindow) {
          appState.jellyfinSetupWindow.destroy();
        }
        appState.jellyfinSetupWindow = null;
        stopJellyfinRemoteSession();
      },
      shouldRestoreWindowsOnActivate: () =>
        appState.overlayRuntimeInitialized && BrowserWindow.getAllWindows().length === 0,
      restoreWindowsOnActivate: () => {
        createMainWindow();
        createInvisibleWindow();
        overlayVisibilityRuntime.updateVisibleOverlayVisibility();
        overlayVisibilityRuntime.updateInvisibleOverlayVisibility();
      },
      shouldQuitOnWindowAllClosed: () => !appState.backgroundMode,
    }),
  }),
);

applyStartupState(appState, startupState);
void refreshAnilistClientSecretState({ force: true });
anilistStateRuntime.refreshRetryQueueState();

function handleCliCommand(args: CliArgs, source: CliCommandSource = 'initial'): void {
  createHandleTexthookerOnlyModeTransitionHandler({
    isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
    setTexthookerOnlyMode: (enabled) => {
      appState.texthookerOnlyMode = enabled;
    },
    commandNeedsOverlayRuntime: (inputArgs) => commandNeedsOverlayRuntime(inputArgs),
    startBackgroundWarmups: () => startBackgroundWarmups(),
    logInfo: (message) => logger.info(message),
  })(args);

  const cliContext = createCliCommandContext({
    getSocketPath: () => appState.mpvSocketPath,
    setSocketPath: (socketPath: string) => {
      appState.mpvSocketPath = socketPath;
    },
    getMpvClient: () => appState.mpvClient,
    showOsd: (text: string) => showMpvOsd(text),
    texthookerService,
    getTexthookerPort: () => appState.texthookerPort,
    setTexthookerPort: (port: number) => {
      appState.texthookerPort = port;
    },
    shouldOpenBrowser: () => getResolvedConfig().texthooker?.openBrowser !== false,
    openExternal: (url: string) => shell.openExternal(url),
    logBrowserOpenError: (url: string, error: unknown) =>
      logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
    isOverlayInitialized: () => appState.overlayRuntimeInitialized,
    initializeOverlay: () => initializeOverlayRuntime(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    toggleInvisibleOverlay: () => toggleInvisibleOverlay(),
    setVisibleOverlay: (visible: boolean) => setVisibleOverlayVisible(visible),
    setInvisibleOverlay: (visible: boolean) => setInvisibleOverlayVisible(visible),
    copyCurrentSubtitle: () => copyCurrentSubtitle(),
    startPendingMultiCopy: (timeoutMs: number) => startPendingMultiCopy(timeoutMs),
    mineSentenceCard: () => mineSentenceCard(),
    startPendingMineSentenceMultiple: (timeoutMs: number) =>
      startPendingMineSentenceMultiple(timeoutMs),
    updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
    refreshKnownWordCache: () => refreshKnownWordCache(),
    triggerFieldGrouping: () => triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    markLastCardAsAudioCard: () => markLastCardAsAudioCard(),
    getAnilistStatus: () => anilistStateRuntime.getStatusSnapshot(),
    clearAnilistToken: () => anilistStateRuntime.clearTokenState(),
    openAnilistSetup: () => openAnilistSetupWindow(),
    openJellyfinSetup: () => openJellyfinSetupWindow(),
    getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
    retryAnilistQueueNow: () => processNextAnilistRetryUpdate(),
    runJellyfinCommand: (argsFromCommand: CliArgs) => runJellyfinCommand(argsFromCommand),
    openYomitanSettings: () => openYomitanSettings(),
    cycleSecondarySubMode: () => cycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
    stopApp: () => app.quit(),
    hasMainWindow: () => Boolean(overlayManager.getMainWindow()),
    getMultiCopyTimeoutMs: () => getConfiguredShortcuts().multiCopyTimeoutMs,
    schedule: (fn: () => void, delayMs: number) => setTimeout(fn, delayMs),
    logInfo: (message: string) => logger.info(message),
    logWarn: (message: string) => logger.warn(message),
    logError: (message: string, err: unknown) => logger.error(message, err),
  });
  handleCliCommandRuntimeServiceWithContext(args, source, cliContext);
}

function handleInitialArgs(): void {
  createHandleInitialArgsHandler({
    getInitialArgs: () => appState.initialArgs,
    isBackgroundMode: () => appState.backgroundMode,
    ensureTray: () => ensureTray(),
    isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
    hasImmersionTracker: () => Boolean(appState.immersionTracker),
    getMpvClient: () => appState.mpvClient,
    logInfo: (message) => logger.info(message),
    handleCliCommand: (args, source) => handleCliCommand(args, source),
  })();
}

function bindMpvClientEventHandlers(mpvClient: MpvIpcClient): void {
  const handleMpvConnectionChange = createHandleMpvConnectionChangeHandler({
    reportJellyfinRemoteStopped: () => {
      void reportJellyfinRemoteStopped();
    },
    hasInitialJellyfinPlayArg: () => Boolean(appState.initialArgs?.jellyfinPlay),
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    isQuitOnDisconnectArmed: () => jellyfinPlayQuitOnDisconnectArmed,
    scheduleQuitCheck: (callback) => {
      setTimeout(callback, 500);
    },
    isMpvConnected: () => Boolean(appState.mpvClient?.connected),
    quitApp: () => app.quit(),
  });
  const handleMpvSubtitleTiming = createHandleMpvSubtitleTimingHandler({
    recordImmersionSubtitleLine: (text, start, end) => {
      appState.immersionTracker?.recordSubtitleLine(text, start, end);
    },
    hasSubtitleTimingTracker: () => Boolean(appState.subtitleTimingTracker),
    recordSubtitleTiming: (text, start, end) => {
      appState.subtitleTimingTracker?.recordSubtitle(text, start, end);
    },
    maybeRunAnilistPostWatchUpdate: () => maybeRunAnilistPostWatchUpdate(),
    logError: (message, error) => logger.error(message, error),
  });
  createBindMpvClientEventHandlers({
    onConnectionChange: (payload) => {
      handleMpvConnectionChange(payload);
    },
    onSubtitleChange: ({ text }) => {
      appState.currentSubText = text;
      broadcastToOverlayWindows('subtitle:set', { text, tokens: null });
      subtitleProcessingController.onSubtitleChange(text);
    },
    onSubtitleAssChange: ({ text }) => {
      appState.currentSubAssText = text;
      broadcastToOverlayWindows('subtitle-ass:set', text);
    },
    onSecondarySubtitleChange: ({ text }) => {
      broadcastToOverlayWindows('secondary-subtitle:set', text);
    },
    onSubtitleTiming: (payload) => {
      handleMpvSubtitleTiming(payload);
    },
    onMediaPathChange: ({ path }) => {
      mediaRuntime.updateCurrentMediaPath(path);
      if (!path) {
        void reportJellyfinRemoteStopped();
      }
      const mediaKey = getCurrentAnilistMediaKey();
      resetAnilistMediaTracking(mediaKey);
      if (mediaKey) {
        void maybeProbeAnilistDuration(mediaKey);
        void ensureAnilistMediaGuess(mediaKey);
      }
      immersionMediaRuntime.syncFromCurrentMediaState();
    },
    onMediaTitleChange: ({ title }) => {
      mediaRuntime.updateCurrentMediaTitle(title);
      anilistCurrentMediaGuess = null;
      anilistCurrentMediaGuessPromise = null;
      appState.immersionTracker?.handleMediaTitleUpdate(title);
      immersionMediaRuntime.syncFromCurrentMediaState();
    },
    onTimePosChange: ({ time }) => {
      appState.immersionTracker?.recordPlaybackPosition(time);
      void reportJellyfinRemoteProgress(false);
    },
    onPauseChange: ({ paused }) => {
      appState.immersionTracker?.recordPauseState(paused);
      void reportJellyfinRemoteProgress(true);
    },
    onSubtitleMetricsChange: ({ patch }) => {
      updateMpvSubtitleRenderMetrics(patch as Partial<MpvSubtitleRenderMetrics>);
    },
    onSecondarySubtitleVisibility: ({ visible }) => {
      appState.previousSecondarySubVisibility = visible;
    },
  })(mpvClient);
}

function createMpvClientRuntimeService(): MpvIpcClient {
  return createMpvClientRuntimeServiceFactory({
    createClient: MpvIpcClient,
    socketPath: appState.mpvSocketPath,
    options: {
      getResolvedConfig: () => getResolvedConfig(),
      autoStartOverlay: appState.autoStartOverlay,
      setOverlayVisible: (visible: boolean) => setOverlayVisible(visible),
      shouldBindVisibleOverlayToMpvSubVisibility: () =>
        configDerivedRuntime.shouldBindVisibleOverlayToMpvSubVisibility(),
      isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      getReconnectTimer: () => appState.reconnectTimer,
      setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
        appState.reconnectTimer = timer;
      },
    },
    bindEventHandlers: (client) => bindMpvClientEventHandlers(client),
  })();
}

const updateMpvSubtitleRenderMetricsRuntime = createUpdateMpvSubtitleRenderMetricsHandler({
  getCurrentMetrics: () => appState.mpvSubtitleRenderMetrics,
  setCurrentMetrics: (metrics) => {
    appState.mpvSubtitleRenderMetrics = metrics;
  },
  applyPatch: (current, patch) => applyMpvSubtitleRenderMetricsPatch(current, patch),
  broadcastMetrics: (metrics) => {
    broadcastToOverlayWindows('mpv-subtitle-render-metrics:set', metrics);
  },
});

function updateMpvSubtitleRenderMetrics(patch: Partial<MpvSubtitleRenderMetrics>): void {
  updateMpvSubtitleRenderMetricsRuntime(patch);
}

async function tokenizeSubtitle(text: string): Promise<SubtitleData> {
  await jlptDictionaryRuntime.ensureJlptDictionaryLookup();
  await frequencyDictionaryRuntime.ensureFrequencyDictionaryLookup();
  return tokenizeSubtitleCore(
    text,
    createTokenizerDepsRuntime({
      getYomitanExt: () => appState.yomitanExt,
      getYomitanParserWindow: () => appState.yomitanParserWindow,
      setYomitanParserWindow: (window) => {
        appState.yomitanParserWindow = window;
      },
      getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
      setYomitanParserReadyPromise: (promise) => {
        appState.yomitanParserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
      setYomitanParserInitPromise: (promise) => {
        appState.yomitanParserInitPromise = promise;
      },
      isKnownWord: (text) =>
        (() => {
          const hit = Boolean(appState.ankiIntegration?.isKnownWord(text));
          appState.immersionTracker?.recordLookup(hit);
          return hit;
        })(),
      getKnownWordMatchMode: () =>
        appState.ankiIntegration?.getKnownWordMatchMode() ??
        getResolvedConfig().ankiConnect.nPlusOne.matchMode,
      getMinSentenceWordsForNPlusOne: () =>
        getResolvedConfig().ankiConnect.nPlusOne.minSentenceWords,
      getJlptLevel: (text) => appState.jlptLevelLookup(text),
      getJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
      getFrequencyDictionaryEnabled: () =>
        getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
      getFrequencyRank: (text) => appState.frequencyRankLookup(text),
      getYomitanGroupDebugEnabled: () => appState.overlayDebugVisualizationEnabled,
      getMecabTokenizer: () => appState.mecabTokenizer,
    }),
  );
}

async function createMecabTokenizerAndCheck(): Promise<void> {
  if (!appState.mecabTokenizer) {
    appState.mecabTokenizer = new MecabTokenizer();
  }
  await appState.mecabTokenizer.checkAvailability();
}

async function prewarmSubtitleDictionaries(): Promise<void> {
  await Promise.all([
    jlptDictionaryRuntime.ensureJlptDictionaryLookup(),
    frequencyDictionaryRuntime.ensureFrequencyDictionaryLookup(),
  ]);
}

const launchBackgroundWarmupTask = createLaunchBackgroundWarmupTaskHandler({
  now: () => Date.now(),
  logDebug: (message) => logger.debug(message),
  logWarn: (message) => logger.warn(message),
});

const startBackgroundWarmups = createStartBackgroundWarmupsHandler({
  getStarted: () => backgroundWarmupsStarted,
  setStarted: (started) => {
    backgroundWarmupsStarted = started;
  },
  isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
  launchTask: (label, task) => launchBackgroundWarmupTask(label, task),
  createMecabTokenizerAndCheck: () => createMecabTokenizerAndCheck(),
  ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded().then(() => {}),
  prewarmSubtitleDictionaries: () => prewarmSubtitleDictionaries(),
  shouldAutoConnectJellyfinRemote: () => getResolvedConfig().jellyfin.remoteControlAutoConnect,
  startJellyfinRemoteSession: () => startJellyfinRemoteSession(),
});

function updateVisibleOverlayBounds(geometry: WindowGeometry): void {
  overlayManager.setOverlayWindowBounds('visible', geometry);
}

function updateInvisibleOverlayBounds(geometry: WindowGeometry): void {
  overlayManager.setOverlayWindowBounds('invisible', geometry);
}

function ensureOverlayWindowLevel(window: BrowserWindow): void {
  ensureOverlayWindowLevelCore(window);
}

function enforceOverlayLayerOrder(): void {
  enforceOverlayLayerOrderCore({
    visibleOverlayVisible: overlayManager.getVisibleOverlayVisible(),
    invisibleOverlayVisible: overlayManager.getInvisibleOverlayVisible(),
    mainWindow: overlayManager.getMainWindow(),
    invisibleWindow: overlayManager.getInvisibleWindow(),
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
  });
}

async function loadYomitanExtension(): Promise<Extension | null> {
  return loadYomitanExtensionCore({
    userDataPath: USER_DATA_PATH,
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    setYomitanParserWindow: (window) => {
      appState.yomitanParserWindow = window;
    },
    setYomitanParserReadyPromise: (promise) => {
      appState.yomitanParserReadyPromise = promise;
    },
    setYomitanParserInitPromise: (promise) => {
      appState.yomitanParserInitPromise = promise;
    },
    setYomitanExtension: (extension) => {
      appState.yomitanExt = extension;
    },
  });
}

async function ensureYomitanExtensionLoaded(): Promise<Extension | null> {
  if (appState.yomitanExt) {
    return appState.yomitanExt;
  }
  if (yomitanLoadInFlight) {
    return yomitanLoadInFlight;
  }

  yomitanLoadInFlight = loadYomitanExtension().finally(() => {
    yomitanLoadInFlight = null;
  });
  return yomitanLoadInFlight;
}

function createOverlayWindow(kind: 'visible' | 'invisible'): BrowserWindow {
  return createOverlayWindowCore(kind, {
    isDev,
    overlayDebugVisualizationEnabled: appState.overlayDebugVisualizationEnabled,
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
    onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    setOverlayDebugVisualizationEnabled: (enabled) => setOverlayDebugVisualizationEnabled(enabled),
    isOverlayVisible: (windowKind) =>
      windowKind === 'visible'
        ? overlayManager.getVisibleOverlayVisible()
        : overlayManager.getInvisibleOverlayVisible(),
    tryHandleOverlayShortcutLocalFallback: (input) =>
      overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(input),
    onWindowClosed: (windowKind) => {
      if (windowKind === 'visible') {
        overlayManager.setMainWindow(null);
      } else {
        overlayManager.setInvisibleWindow(null);
      }
    },
  });
}

function createMainWindow(): BrowserWindow {
  const window = createOverlayWindow('visible');
  overlayManager.setMainWindow(window);
  return window;
}
function createInvisibleWindow(): BrowserWindow {
  const window = createOverlayWindow('invisible');
  overlayManager.setInvisibleWindow(window);
  return window;
}

function resolveTrayIconPath(): string | null {
  const iconNames =
    process.platform === 'darwin'
      ? ['SubMinerTemplate.png', 'SubMinerTemplate@2x.png', 'SubMiner.png']
      : ['SubMiner.png'];

  const baseDirs = [
    path.join(process.resourcesPath, 'assets'),
    path.join(app.getAppPath(), 'assets'),
    path.join(__dirname, '..', 'assets'),
    path.join(__dirname, '..', '..', 'assets'),
  ];

  const candidates = baseDirs.flatMap((baseDir) =>
    iconNames.map((iconName) => path.join(baseDir, iconName)),
  );
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return null;
}

function buildTrayMenu(): Menu {
  return Menu.buildFromTemplate([
    {
      label: 'Open Overlay',
      click: () => {
        if (!appState.overlayRuntimeInitialized) {
          initializeOverlayRuntime();
        }
        setVisibleOverlayVisible(true);
      },
    },
    {
      label: 'Open Yomitan Settings',
      click: () => {
        openYomitanSettings();
      },
    },
    {
      label: 'Open Runtime Options',
      click: () => {
        if (!appState.overlayRuntimeInitialized) {
          initializeOverlayRuntime();
        }
        openRuntimeOptionsPalette();
      },
    },
    {
      label: 'Configure Jellyfin',
      click: () => {
        openJellyfinSetupWindow();
      },
    },
    {
      label: 'Configure AniList',
      click: () => {
        openAnilistSetupWindow();
      },
    },
    { type: 'separator' },
    {
      label: 'Quit',
      click: () => {
        app.quit();
      },
    },
  ]);
}

function ensureTray(): void {
  if (appTray) {
    appTray.setContextMenu(buildTrayMenu());
    return;
  }

  const iconPath = resolveTrayIconPath();
  let trayIcon = iconPath ? nativeImage.createFromPath(iconPath) : nativeImage.createEmpty();
  if (trayIcon.isEmpty()) {
    logger.warn('Tray icon asset not found; using empty icon placeholder.');
  }
  if (process.platform === 'darwin' && !trayIcon.isEmpty()) {
    // macOS status bar expects a small monochrome-like template icon.
    // Feeding the full-size app icon can produce oversized/non-interactive items.
    trayIcon = trayIcon.resize({ width: 18, height: 18, quality: 'best' });
    trayIcon.setTemplateImage(true);
  }
  if (process.platform === 'linux' && !trayIcon.isEmpty()) {
    trayIcon = trayIcon.resize({ width: 20, height: 20 });
  }

  appTray = new Tray(trayIcon);
  appTray.setToolTip(TRAY_TOOLTIP);
  appTray.setContextMenu(buildTrayMenu());
  appTray.on('click', () => {
    if (!appState.overlayRuntimeInitialized) {
      initializeOverlayRuntime();
    }
    setVisibleOverlayVisible(true);
  });
}

function destroyTray(): void {
  if (!appTray) {
    return;
  }
  appTray.destroy();
  appTray = null;
}

function initializeOverlayRuntime(): void {
  if (appState.overlayRuntimeInitialized) {
    return;
  }
  const result = initializeOverlayRuntimeCore({
    backendOverride: appState.backendOverride,
    getInitialInvisibleOverlayVisibility: () =>
      configDerivedRuntime.getInitialInvisibleOverlayVisibility(),
    createMainWindow: () => {
      createMainWindow();
    },
    createInvisibleWindow: () => {
      createInvisibleWindow();
    },
    registerGlobalShortcuts: () => {
      registerGlobalShortcuts();
    },
    updateVisibleOverlayBounds: (geometry) => {
      updateVisibleOverlayBounds(geometry);
    },
    updateInvisibleOverlayBounds: (geometry) => {
      updateInvisibleOverlayBounds(geometry);
    },
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    isInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
    updateVisibleOverlayVisibility: () => {
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
    },
    updateInvisibleOverlayVisibility: () => {
      overlayVisibilityRuntime.updateInvisibleOverlayVisibility();
    },
    getOverlayWindows: () => getOverlayWindows(),
    syncOverlayShortcuts: () => overlayShortcutsRuntime.syncOverlayShortcuts(),
    setWindowTracker: (tracker) => {
      appState.windowTracker = tracker;
    },
    getResolvedConfig: () => getResolvedConfig(),
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getMpvClient: () => appState.mpvClient,
    getMpvSocketPath: () => appState.mpvSocketPath,
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
    setAnkiIntegration: (integration) => {
      appState.ankiIntegration = integration as AnkiIntegration | null;
    },
    showDesktopNotification,
    createFieldGroupingCallback: () => createFieldGroupingCallback(),
    getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
  });
  overlayManager.setInvisibleOverlayVisible(result.invisibleOverlayVisible);
  appState.overlayRuntimeInitialized = true;
  startBackgroundWarmups();
}

function openYomitanSettings(): void {
  void (async () => {
    const extension = await ensureYomitanExtensionLoaded();
    if (!extension) {
      logger.warn('Unable to open Yomitan settings: extension failed to load.');
      return;
    }

    openYomitanSettingsWindow({
      yomitanExt: extension,
      getExistingWindow: () => appState.yomitanSettingsWindow,
      setWindow: (window: BrowserWindow | null) => {
        appState.yomitanSettingsWindow = window;
      },
    });
  })().catch((error) => {
    logger.error('Failed to open Yomitan settings window.', error);
  });
}
function registerGlobalShortcuts(): void {
  registerGlobalShortcutsCore({
    shortcuts: getConfiguredShortcuts(),
    onToggleVisibleOverlay: () => toggleVisibleOverlay(),
    onToggleInvisibleOverlay: () => toggleInvisibleOverlay(),
    onOpenYomitanSettings: () => openYomitanSettings(),
    isDev,
    getMainWindow: () => overlayManager.getMainWindow(),
  });
}

function refreshGlobalAndOverlayShortcuts(): void {
  globalShortcut.unregisterAll();
  registerGlobalShortcuts();
  syncOverlayShortcuts();
}

function getConfiguredShortcuts() {
  return resolveConfiguredShortcuts(getResolvedConfig(), DEFAULT_CONFIG);
}

function cycleSecondarySubMode(): void {
  cycleSecondarySubModeCore({
    getSecondarySubMode: () => appState.secondarySubMode,
    setSecondarySubMode: (mode: SecondarySubMode) => {
      appState.secondarySubMode = mode;
    },
    getLastSecondarySubToggleAtMs: () => appState.lastSecondarySubToggleAtMs,
    setLastSecondarySubToggleAtMs: (timestampMs: number) => {
      appState.lastSecondarySubToggleAtMs = timestampMs;
    },
    broadcastSecondarySubMode: (mode: SecondarySubMode) => {
      broadcastToOverlayWindows('secondary-subtitle:mode', mode);
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
  });
}

function showMpvOsd(text: string): void {
  appendToMpvLog(`[OSD] ${text}`);
  showMpvOsdRuntime(appState.mpvClient, text, (line) => {
    logger.info(line);
  });
}

function appendToMpvLog(message: string): void {
  try {
    fs.mkdirSync(path.dirname(DEFAULT_MPV_LOG_PATH), { recursive: true });
    fs.appendFileSync(DEFAULT_MPV_LOG_PATH, `[${new Date().toISOString()}] ${message}\n`, {
      encoding: 'utf8',
    });
  } catch {
    // best-effort logging
  }
}

const numericShortcutRuntime = createNumericShortcutRuntime({
  globalShortcut,
  showMpvOsd: (text) => showMpvOsd(text),
  setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
  clearTimer: (timer) => clearTimeout(timer),
});
const multiCopySession = numericShortcutRuntime.createSession();
const mineSentenceSession = numericShortcutRuntime.createSession();

async function triggerSubsyncFromConfig(): Promise<void> {
  await subsyncRuntime.triggerFromConfig();
}

function cancelPendingMultiCopy(): void {
  multiCopySession.cancel();
}

function startPendingMultiCopy(timeoutMs: number): void {
  multiCopySession.start({
    timeoutMs,
    onDigit: (count) => handleMultiCopyDigit(count),
    messages: {
      prompt: 'Copy how many lines? Press 1-9 (Esc to cancel)',
      timeout: 'Copy timeout',
      cancelled: 'Cancelled',
    },
  });
}

function handleMultiCopyDigit(count: number): void {
  handleMultiCopyDigitCore(count, {
    subtitleTimingTracker: appState.subtitleTimingTracker,
    writeClipboardText: (text) => clipboard.writeText(text),
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

function copyCurrentSubtitle(): void {
  copyCurrentSubtitleCore({
    subtitleTimingTracker: appState.subtitleTimingTracker,
    writeClipboardText: (text) => clipboard.writeText(text),
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function updateLastCardFromClipboard(): Promise<void> {
  await updateLastCardFromClipboardCore({
    ankiIntegration: appState.ankiIntegration,
    readClipboardText: () => clipboard.readText(),
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function refreshKnownWordCache(): Promise<void> {
  if (!appState.ankiIntegration) {
    throw new Error('AnkiConnect integration not enabled');
  }

  await appState.ankiIntegration.refreshKnownWordCache();
}

async function triggerFieldGrouping(): Promise<void> {
  await triggerFieldGroupingCore({
    ankiIntegration: appState.ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function markLastCardAsAudioCard(): Promise<void> {
  await markLastCardAsAudioCardCore({
    ankiIntegration: appState.ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function mineSentenceCard(): Promise<void> {
  const created = await mineSentenceCardCore({
    ankiIntegration: appState.ankiIntegration,
    mpvClient: appState.mpvClient,
    showMpvOsd: (text) => showMpvOsd(text),
  });
  if (created) {
    appState.immersionTracker?.recordCardsMined(1);
  }
}

function cancelPendingMineSentenceMultiple(): void {
  mineSentenceSession.cancel();
}

function startPendingMineSentenceMultiple(timeoutMs: number): void {
  mineSentenceSession.start({
    timeoutMs,
    onDigit: (count) => handleMineSentenceDigit(count),
    messages: {
      prompt: 'Mine how many lines? Press 1-9 (Esc to cancel)',
      timeout: 'Mine sentence timeout',
      cancelled: 'Cancelled',
    },
  });
}

function handleMineSentenceDigit(count: number): void {
  handleMineSentenceDigitCore(count, {
    subtitleTimingTracker: appState.subtitleTimingTracker,
    ankiIntegration: appState.ankiIntegration,
    getCurrentSecondarySubText: () => appState.mpvClient?.currentSecondarySubText || undefined,
    showMpvOsd: (text) => showMpvOsd(text),
    logError: (message, err) => {
      logger.error(message, err);
    },
    onCardsMined: (cards) => {
      appState.immersionTracker?.recordCardsMined(cards);
    },
  });
}

function registerOverlayShortcuts(): void {
  overlayShortcutsRuntime.registerOverlayShortcuts();
}

function unregisterOverlayShortcuts(): void {
  overlayShortcutsRuntime.unregisterOverlayShortcuts();
}

function syncOverlayShortcuts(): void {
  overlayShortcutsRuntime.syncOverlayShortcuts();
}
function refreshOverlayShortcuts(): void {
  overlayShortcutsRuntime.refreshOverlayShortcuts();
}

function setVisibleOverlayVisible(visible: boolean): void {
  setVisibleOverlayVisibleCore({
    visible,
    setVisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setVisibleOverlayVisible(nextVisible);
    },
    updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () =>
      overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      overlayVisibilityRuntime.syncInvisibleOverlayMousePassthrough(),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      configDerivedRuntime.shouldBindVisibleOverlayToMpvSubVisibility(),
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    setMpvSubVisibility: (mpvSubVisible) => {
      setMpvSubVisibilityRuntime(appState.mpvClient, mpvSubVisible);
    },
  });
}

function setInvisibleOverlayVisible(visible: boolean): void {
  setInvisibleOverlayVisibleCore({
    visible,
    setInvisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setInvisibleOverlayVisible(nextVisible);
    },
    updateInvisibleOverlayVisibility: () =>
      overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      overlayVisibilityRuntime.syncInvisibleOverlayMousePassthrough(),
  });
}

function toggleVisibleOverlay(): void {
  setVisibleOverlayVisible(!overlayManager.getVisibleOverlayVisible());
}
function toggleInvisibleOverlay(): void {
  setInvisibleOverlayVisible(!overlayManager.getInvisibleOverlayVisible());
}
function setOverlayVisible(visible: boolean): void {
  setVisibleOverlayVisible(visible);
}
function toggleOverlay(): void {
  toggleVisibleOverlay();
}
function handleOverlayModalClosed(modal: OverlayHostedModal): void {
  overlayModalRuntime.handleOverlayModalClosed(modal);
}

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcRuntime(command, {
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    cycleRuntimeOption: (id, direction) => {
      if (!appState.runtimeOptionsManager) {
        return { ok: false, error: 'Runtime options manager unavailable' };
      }
      return applyRuntimeOptionResultRuntime(
        appState.runtimeOptionsManager.cycleOption(id, direction),
        (text) => showMpvOsd(text),
      );
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
    replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(appState.mpvClient),
    playNextSubtitle: () => playNextSubtitleRuntime(appState.mpvClient),
    sendMpvCommand: (rawCommand: (string | number)[]) =>
      sendMpvCommandRuntime(appState.mpvClient, rawCommand),
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
  });
}

async function runSubsyncManualFromIpc(request: SubsyncManualRunRequest): Promise<SubsyncResult> {
  return subsyncRuntime.runManualFromIpc(request);
}

function appendClipboardVideoToQueue(): { ok: boolean; message: string } {
  return appendClipboardVideoToQueueRuntime({
    getMpvClient: () => appState.mpvClient,
    readClipboardText: () => clipboard.readText(),
    showMpvOsd: (text) => showMpvOsd(text),
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
  });
}

registerIpcRuntimeServices({
  runtimeOptions: {
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
    showMpvOsd: (text: string) => showMpvOsd(text),
  },
  mainDeps: {
    getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisibility: () => overlayManager.getVisibleOverlayVisible(),
    getInvisibleOverlayVisibility: () => overlayManager.getInvisibleOverlayVisible(),
    focusMainWindow: () => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      if (!mainWindow.isFocused()) {
        mainWindow.focus();
      }
    },
    onOverlayModalClosed: (modal: string) => {
      handleOverlayModalClosed(modal as OverlayHostedModal);
    },
    openYomitanSettings: () => openYomitanSettings(),
    quitApp: () => app.quit(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    tokenizeCurrentSubtitle: () => tokenizeSubtitle(appState.currentSubText),
    getCurrentSubtitleRaw: () => appState.currentSubText,
    getCurrentSubtitleAss: () => appState.currentSubAssText,
    getMpvSubtitleRenderMetrics: () => appState.mpvSubtitleRenderMetrics,
    getSubtitlePosition: () => loadSubtitlePosition(),
    getSubtitleStyle: () => {
      const resolvedConfig = getResolvedConfig();
      return resolveSubtitleStyleForRenderer(resolvedConfig);
    },
    saveSubtitlePosition: (position: unknown) => saveSubtitlePosition(position as SubtitlePosition),
    getMecabTokenizer: () => appState.mecabTokenizer,
    handleMpvCommand: (command: (string | number)[]) => handleMpvCommandFromIpc(command),
    getKeybindings: () => appState.keybindings,
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getSecondarySubMode: () => appState.secondarySubMode,
    getMpvClient: () => appState.mpvClient,
    runSubsyncManual: (request: unknown) =>
      runSubsyncManualFromIpc(request as SubsyncManualRunRequest),
    getAnkiConnectStatus: () => appState.ankiIntegration !== null,
    getRuntimeOptions: () => getRuntimeOptionsState(),
    reportOverlayContentBounds: (payload: unknown) => {
      overlayContentMeasurementStore.report(payload);
    },
    getAnilistStatus: () => anilistStateRuntime.getStatusSnapshot(),
    clearAnilistToken: () => anilistStateRuntime.clearTokenState(),
    openAnilistSetup: () => openAnilistSetupWindow(),
    getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
    retryAnilistQueueNow: () => processNextAnilistRetryUpdate(),
    appendClipboardVideoToQueue: () => appendClipboardVideoToQueue(),
  },
  ankiJimakuDeps: createAnkiJimakuIpcRuntimeServiceDeps({
    patchAnkiConnectEnabled: (enabled: boolean) => {
      configService.patchRawConfig({ ankiConnect: { enabled } });
    },
    getResolvedConfig: () => getResolvedConfig(),
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getMpvClient: () => appState.mpvClient,
    getAnkiIntegration: () => appState.ankiIntegration,
    setAnkiIntegration: (integration: AnkiIntegration | null) => {
      appState.ankiIntegration = integration;
    },
    getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
    showDesktopNotification,
    createFieldGroupingCallback: () => createFieldGroupingCallback(),
    broadcastRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    getFieldGroupingResolver: () => getFieldGroupingResolver(),
    setFieldGroupingResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) =>
      setFieldGroupingResolver(resolver),
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
    downloadToFile: (url: string, destPath: string, headers: Record<string, string>) =>
      downloadToFile(url, destPath, headers),
  }),
});
