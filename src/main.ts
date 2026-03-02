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
  Menu,
  nativeImage,
  Tray,
  dialog,
  screen,
} from 'electron';

function getPasswordStoreArg(argv: string[]): string | null {
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg?.startsWith('--password-store')) {
      continue;
    }

    if (arg === '--password-store') {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        return value;
      }
      return null;
    }

    const [prefix, value] = arg.split('=', 2);
    if (prefix === '--password-store' && value && value.trim().length > 0) {
      return value.trim();
    }
  }
  return null;
}

function normalizePasswordStoreArg(value: string): string {
  const normalized = value.trim();
  if (normalized.toLowerCase() === 'gnome') {
    return 'gnome-libsecret';
  }
  return normalized;
}

function getDefaultPasswordStore(): string {
  return 'gnome-libsecret';
}

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
import { MecabTokenizer } from './mecab-tokenizer';
import type {
  JimakuApiResponse,
  KikuFieldGroupingChoice,
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
  RuntimeOptionState,
  SecondarySubMode,
  SubtitleData,
  SubtitlePosition,
  SubsyncManualRunRequest,
  SubsyncResult,
  WindowGeometry,
} from './types';
import { AnkiIntegration } from './anki-integration';
import { SubtitleTimingTracker } from './subtitle-timing-tracker';
import { RuntimeOptionsManager } from './runtime-options';
import { downloadToFile, isRemoteMediaPath, parseMediaInfo } from './jimaku/utils';
import { createLogger, setLogLevel, type LogLevelSource } from './logger';
import {
  commandNeedsOverlayRuntime,
  parseArgs,
  shouldRunSettingsOnlyStartup,
  shouldStartApp,
} from './cli/args';
import type { CliArgs, CliCommandSource } from './cli/args';
import { printHelp } from './cli/help';
import {
  buildConfigParseErrorDetails,
  buildConfigWarningDialogDetails,
  buildConfigWarningNotificationBody,
  failStartupFromConfig,
} from './main/config-validation';
import {
  buildAnilistAttemptKey,
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  createAnilistStateRuntime,
  createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler,
  createBuildGetCurrentAnilistMediaKeyMainDepsHandler,
  createBuildMaybeProbeAnilistDurationMainDepsHandler,
  createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler,
  createBuildOpenAnilistSetupWindowMainDepsHandler,
  createBuildProcessNextAnilistRetryUpdateMainDepsHandler,
  createBuildRefreshAnilistClientSecretStateMainDepsHandler,
  createBuildResetAnilistMediaGuessStateMainDepsHandler,
  createBuildResetAnilistMediaTrackingMainDepsHandler,
  createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler,
  createEnsureAnilistMediaGuessHandler,
  createGetAnilistMediaGuessRuntimeStateHandler,
  createGetCurrentAnilistMediaKeyHandler,
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createMaybeProbeAnilistDurationHandler,
  createMaybeRunAnilistPostWatchUpdateHandler,
  createOpenAnilistSetupWindowHandler,
  createProcessNextAnilistRetryUpdateHandler,
  createRefreshAnilistClientSecretStateHandler,
  createResetAnilistMediaGuessStateHandler,
  createResetAnilistMediaTrackingHandler,
  createSetAnilistMediaGuessRuntimeStateHandler,
  findAnilistSetupDeepLinkArgvUrl,
  isAnilistTrackingEnabled,
  loadAnilistManualTokenEntry,
  loadAnilistSetupFallback,
  openAnilistSetupInBrowser,
  rememberAnilistAttemptedUpdateKey,
} from './main/runtime/domains/anilist';
import {
  createApplyJellyfinMpvDefaultsHandler,
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler,
  createBuildGetDefaultSocketPathMainDepsHandler,
  createEnsureMpvConnectedForJellyfinPlaybackHandler,
  createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler,
  createGetDefaultSocketPathHandler,
  createGetJellyfinClientInfoHandler,
  createBuildGetJellyfinClientInfoMainDepsHandler,
  createHandleJellyfinAuthCommands,
  createBuildHandleJellyfinAuthCommandsMainDepsHandler,
  createHandleJellyfinListCommands,
  createBuildHandleJellyfinListCommandsMainDepsHandler,
  createHandleJellyfinPlayCommand,
  createBuildHandleJellyfinPlayCommandMainDepsHandler,
  createHandleJellyfinRemoteAnnounceCommand,
  createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler,
  createHandleJellyfinRemotePlay,
  createBuildHandleJellyfinRemotePlayMainDepsHandler,
  createHandleJellyfinRemotePlaystate,
  createBuildHandleJellyfinRemotePlaystateMainDepsHandler,
  createHandleJellyfinRemoteGeneralCommand,
  createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler,
  createLaunchMpvIdleForJellyfinPlaybackHandler,
  createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler,
  createPlayJellyfinItemInMpvHandler,
  createBuildPlayJellyfinItemInMpvMainDepsHandler,
  createPreloadJellyfinExternalSubtitlesHandler,
  createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler,
  createReportJellyfinRemoteProgressHandler,
  createBuildReportJellyfinRemoteProgressMainDepsHandler,
  createReportJellyfinRemoteStoppedHandler,
  createBuildReportJellyfinRemoteStoppedMainDepsHandler,
  createStartJellyfinRemoteSessionHandler,
  createBuildStartJellyfinRemoteSessionMainDepsHandler,
  createStopJellyfinRemoteSessionHandler,
  createBuildStopJellyfinRemoteSessionMainDepsHandler,
  createRunJellyfinCommandHandler,
  createBuildRunJellyfinCommandMainDepsHandler,
  createWaitForMpvConnectedHandler,
  createBuildWaitForMpvConnectedMainDepsHandler,
  createOpenJellyfinSetupWindowHandler,
  createBuildOpenJellyfinSetupWindowMainDepsHandler,
  createGetResolvedJellyfinConfigHandler,
  createBuildGetResolvedJellyfinConfigMainDepsHandler,
  parseJellyfinSetupSubmissionUrl,
  buildJellyfinSetupFormHtml,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
} from './main/runtime/domains/jellyfin';
import type { ActiveJellyfinRemotePlaybackState } from './main/runtime/domains/jellyfin';
import { getConfiguredJellyfinSession } from './main/runtime/domains/jellyfin';
import {
  createBuildConfigHotReloadMessageMainDepsHandler,
  createBuildConfigHotReloadAppliedMainDepsHandler,
  createBuildConfigHotReloadRuntimeMainDepsHandler,
  createBuildWatchConfigPathMainDepsHandler,
  createWatchConfigPathHandler,
  createBuildOverlayContentMeasurementStoreMainDepsHandler,
  createBuildOverlayModalRuntimeMainDepsHandler,
  createBuildAppendClipboardVideoToQueueMainDepsHandler,
  createBuildHandleOverlayModalClosedMainDepsHandler,
  createBuildLoadSubtitlePositionMainDepsHandler,
  createBuildSaveSubtitlePositionMainDepsHandler,
  createBuildFieldGroupingOverlayMainDepsHandler,
  createBuildGetFieldGroupingResolverMainDepsHandler,
  createBuildSetFieldGroupingResolverMainDepsHandler,
  createBuildOverlayVisibilityRuntimeMainDepsHandler,
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler,
  createBuildGetRuntimeOptionsStateMainDepsHandler,
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler,
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler,
  createBuildSendToActiveOverlayWindowMainDepsHandler,
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler,
  createBuildEnforceOverlayLayerOrderMainDepsHandler,
  createBuildEnsureOverlayWindowLevelMainDepsHandler,
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler,
  createOverlayWindowRuntimeHandlers,
  createOverlayRuntimeBootstrapHandlers,
  createTrayRuntimeHandlers,
  createOverlayVisibilityRuntime,
  createBroadcastRuntimeOptionsChangedHandler,
  createGetRuntimeOptionsStateHandler,
  createGetFieldGroupingResolverHandler,
  createSetFieldGroupingResolverHandler,
  createOpenRuntimeOptionsPaletteHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
  createSendToActiveOverlayWindowHandler,
  createSetOverlayDebugVisualizationEnabledHandler,
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateVisibleOverlayBoundsHandler,
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
  createAppendClipboardVideoToQueueHandler,
  createHandleOverlayModalClosedHandler,
  createConfigHotReloadMessageHandler,
  createConfigHotReloadAppliedHandler,
  buildTrayMenuTemplateRuntime,
  resolveTrayIconPathRuntime,
  createYomitanExtensionRuntime,
  createYomitanSettingsRuntime,
  buildRestartRequiredConfigMessage,
  resolveSubtitleStyleForRenderer,
} from './main/runtime/domains/overlay';
import {
  createBuildAnilistStateRuntimeMainDepsHandler,
  createBuildConfigDerivedRuntimeMainDepsHandler,
  createBuildImmersionMediaRuntimeMainDepsHandler,
  createBuildMainSubsyncRuntimeMainDepsHandler,
  createBuildSubtitleProcessingControllerMainDepsHandler,
  createBuildMediaRuntimeMainDepsHandler,
  createBuildDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRuntimeMainDepsHandler,
  createBuildJlptDictionaryRuntimeMainDepsHandler,
  createImmersionMediaRuntime,
  createConfigDerivedRuntime,
  appendClipboardVideoToQueueRuntime,
  createMainSubsyncRuntime,
  createLaunchBackgroundWarmupTaskHandler,
  createStartBackgroundWarmupsHandler,
  createBuildLaunchBackgroundWarmupTaskMainDepsHandler,
  createBuildStartBackgroundWarmupsMainDepsHandler,
} from './main/runtime/domains/startup';
import {
  createBuildBindMpvMainEventHandlersMainDepsHandler,
  createBuildMpvClientRuntimeServiceFactoryDepsHandler,
  createMpvClientRuntimeServiceFactory,
  createBindMpvMainEventHandlersHandler,
  createBuildTokenizerDepsMainHandler,
  createCreateMecabTokenizerAndCheckMainHandler,
  createPrewarmSubtitleDictionariesMainHandler,
  createUpdateMpvSubtitleRenderMetricsHandler,
  createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler,
  createMpvOsdRuntimeHandlers,
  createCycleSecondarySubModeRuntimeHandler,
} from './main/runtime/domains/mpv';
import type { MpvClientRuntimeServiceOptions } from './main/runtime/domains/mpv';
import {
  createBuildCopyCurrentSubtitleMainDepsHandler,
  createBuildHandleMineSentenceDigitMainDepsHandler,
  createBuildHandleMultiCopyDigitMainDepsHandler,
  createBuildMarkLastCardAsAudioCardMainDepsHandler,
  createBuildMineSentenceCardMainDepsHandler,
  createBuildRefreshKnownWordCacheMainDepsHandler,
  createBuildTriggerFieldGroupingMainDepsHandler,
  createBuildUpdateLastCardFromClipboardMainDepsHandler,
  createMarkLastCardAsAudioCardHandler,
  createMineSentenceCardHandler,
  createRefreshKnownWordCacheHandler,
  createTriggerFieldGroupingHandler,
  createUpdateLastCardFromClipboardHandler,
  createCopyCurrentSubtitleHandler,
  createHandleMineSentenceDigitHandler,
  createHandleMultiCopyDigitHandler,
} from './main/runtime/domains/mining';
import {
  createCliCommandContextFactory,
  createInitialArgsRuntimeHandler,
  createCliCommandRuntimeHandler,
} from './main/runtime/domains/ipc';
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
  generateDefaultConfigFile,
  resolveConfiguredShortcuts,
  resolveKeybindings,
  showDesktopNotification,
} from './core/utils';
import {
  ImmersionTrackerService,
  JellyfinRemoteSessionService,
  MpvIpcClient,
  SubtitleWebSocket,
  Texthooker,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
  applyMpvSubtitleRenderMetricsPatch,
  authenticateWithPasswordRuntime,
  broadcastRuntimeOptionsChangedRuntime,
  copyCurrentSubtitle as copyCurrentSubtitleCore,
  createConfigHotReloadRuntime,
  createDiscordPresenceService,
  createShiftSubtitleDelayToAdjacentCueHandler,
  createFieldGroupingOverlayRuntime,
  createOverlayContentMeasurementStore,
  createOverlayManager,
  createOverlayWindow as createOverlayWindowCore,
  createSubtitleProcessingController,
  createTokenizerDepsRuntime,
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  enforceOverlayLayerOrder as enforceOverlayLayerOrderCore,
  ensureOverlayWindowLevel as ensureOverlayWindowLevelCore,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  hasMpvWebsocketPlugin,
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  jellyfinTicksToSecondsRuntime,
  listJellyfinItemsRuntime,
  listJellyfinLibrariesRuntime,
  listJellyfinSubtitleTracksRuntime,
  loadSubtitlePosition as loadSubtitlePositionCore,
  loadYomitanExtension as loadYomitanExtensionCore,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  mineSentenceCard as mineSentenceCardCore,
  openYomitanSettingsWindow,
  playNextSubtitleRuntime,
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  replayCurrentSubtitleRuntime,
  resolveJellyfinPlaybackPlanRuntime,
  runStartupBootstrapRuntime,
  saveSubtitlePosition as saveSubtitlePositionCore,
  clearYomitanParserCachesForWindow,
  syncYomitanDefaultAnkiServer as syncYomitanDefaultAnkiServerCore,
  sendMpvCommandRuntime,
  setMpvSubVisibilityRuntime,
  setOverlayDebugVisualizationEnabledRuntime,
  syncOverlayWindowLayer,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  showMpvOsdRuntime,
  tokenizeSubtitle as tokenizeSubtitleCore,
  triggerFieldGrouping as triggerFieldGroupingCore,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
} from './core/services';
import { createImmersionTrackerStartupHandler } from './main/runtime/immersion-startup';
import { createBuildImmersionTrackerStartupMainDepsHandler } from './main/runtime/immersion-startup-main-deps';
import { createAnilistUpdateQueue } from './core/services/anilist/anilist-update-queue';
import {
  guessAnilistMediaInfo,
  updateAnilistPostWatchProgress,
} from './core/services/anilist/anilist-updater';
import { createJellyfinTokenStore } from './core/services/jellyfin-token-store';
import { applyRuntimeOptionResultRuntime } from './core/services/runtime-options-ipc';
import { createAnilistTokenStore } from './core/services/anilist/anilist-token-store';
import { createBuildOverlayShortcutsRuntimeMainDepsHandler } from './main/runtime/domains/shortcuts';
import { createMainRuntimeRegistry } from './main/runtime/registry';
import {
  createEnsureOverlayMpvSubtitlesHiddenHandler,
  createRestoreOverlayMpvSubtitlesHandler,
} from './main/runtime/overlay-mpv-sub-visibility';
import {
  composeAnilistSetupHandlers,
  composeAnilistTrackingHandlers,
  composeAppReadyRuntime,
  composeIpcRuntimeHandlers,
  composeJellyfinRuntimeHandlers,
  composeMpvRuntimeHandlers,
  composeShortcutRuntimes,
  composeStartupLifecycleHandlers,
} from './main/runtime/composers';
import { createStartupBootstrapRuntimeDeps } from './main/startup';
import { createAppLifecycleRuntimeRunner } from './main/startup-lifecycle';
import { handleMpvCommandFromIpcRuntime } from './main/ipc-mpv-command';
import { registerIpcRuntimeServices } from './main/ipc-runtime';
import { createAnkiJimakuIpcRuntimeServiceDeps } from './main/dependencies';
import { handleCliCommandRuntimeServiceWithContext } from './main/cli-runtime';
import { createOverlayModalRuntimeService, type OverlayHostedModal } from './main/overlay-runtime';
import { createOverlayShortcutsRuntimeService } from './main/overlay-shortcuts-runtime';
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from './main/frequency-dictionary-runtime';
import {
  createJlptDictionaryRuntimeService,
  getJlptDictionarySearchPaths,
} from './main/jlpt-runtime';
import { createMediaRuntimeService } from './main/media-runtime';
import { createOverlayVisibilityRuntimeService } from './main/overlay-visibility-runtime';
import {
  type AnilistMediaGuessRuntimeState,
  type AppState,
  type StartupState,
  applyStartupState,
  createAppState,
  createInitialAnilistMediaGuessRuntimeState,
  createInitialAnilistUpdateInFlightState,
  transitionAnilistClientSecretState,
  transitionAnilistMediaGuessRuntimeState,
  transitionAnilistRetryQueueLastAttemptAt,
  transitionAnilistRetryQueueLastError,
  transitionAnilistRetryQueueState,
  transitionAnilistUpdateInFlightState,
} from './main/state';
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

if (process.platform === 'linux') {
  app.commandLine.appendSwitch('enable-features', 'GlobalShortcutsPortal');
  const passwordStore = normalizePasswordStoreArg(
    getPasswordStoreArg(process.argv) ?? getDefaultPasswordStore(),
  );
  app.commandLine.appendSwitch('password-store', passwordStore);
  console.debug(`[main] Applied --password-store ${passwordStore}`);
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
const JELLYFIN_TOKEN_STORE_FILE = 'jellyfin-token-store.json';
const ANILIST_RETRY_QUEUE_FILE = 'anilist-retry-queue.json';
const TRAY_TOOLTIP = 'SubMiner';

let anilistMediaGuessRuntimeState: AnilistMediaGuessRuntimeState =
  createInitialAnilistMediaGuessRuntimeState();
let anilistUpdateInFlightState = createInitialAnilistUpdateInFlightState();
const anilistAttemptedUpdateKeys = new Set<string>();
let anilistCachedAccessToken: string | null = null;
let jellyfinPlayQuitOnDisconnectArmed = false;
const JELLYFIN_LANG_PREF = 'ja,jp,jpn,japanese,en,eng,english,enUS,en-US';
const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS = 3000;
const DISCORD_PRESENCE_APP_ID = '1475264834730856619';
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
let notifyAnilistTokenStoreWarning: (message: string) => void = () => {};

const buildApplyJellyfinMpvDefaultsMainDepsHandler =
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler({
    sendMpvCommandRuntime: (client, command) => sendMpvCommandRuntime(client, command),
    jellyfinLangPref: JELLYFIN_LANG_PREF,
  });
const applyJellyfinMpvDefaultsMainDeps = buildApplyJellyfinMpvDefaultsMainDepsHandler();
const applyJellyfinMpvDefaultsHandler = createApplyJellyfinMpvDefaultsHandler(
  applyJellyfinMpvDefaultsMainDeps,
);

function applyJellyfinMpvDefaults(
  client: Parameters<typeof applyJellyfinMpvDefaultsHandler>[0],
): void {
  applyJellyfinMpvDefaultsHandler(client);
}

const CONFIG_DIR = resolveConfigDir({
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  homeDir: os.homedir(),
  existsSync: fs.existsSync,
});
const USER_DATA_PATH = CONFIG_DIR;
const DEFAULT_MPV_LOG_PATH = process.env.SUBMINER_MPV_LOG?.trim() || DEFAULT_MPV_LOG_FILE;
const DEFAULT_IMMERSION_DB_PATH = path.join(USER_DATA_PATH, 'immersion.sqlite');
const configService = (() => {
  try {
    return new ConfigService(CONFIG_DIR);
  } catch (error) {
    if (error instanceof ConfigStartupParseError) {
      failStartupFromConfig(
        'SubMiner config parse error',
        buildConfigParseErrorDetails(error.path, error.parseError),
        {
          logError: (details) => console.error(details),
          showErrorBox: (title, details) => dialog.showErrorBox(title, details),
          quit: () => app.quit(),
        },
      );
    }
    throw error;
  }
})();
const anilistTokenStore = createAnilistTokenStore(
  path.join(USER_DATA_PATH, ANILIST_TOKEN_STORE_FILE),
  {
    info: (message: string) => console.info(message),
    warn: (message: string, details?: unknown) => console.warn(message, details),
    error: (message: string, details?: unknown) => console.error(message, details),
    warnUser: (message: string) => notifyAnilistTokenStoreWarning(message),
  },
);
const jellyfinTokenStore = createJellyfinTokenStore(
  path.join(USER_DATA_PATH, JELLYFIN_TOKEN_STORE_FILE),
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
const runtimeRegistry = createMainRuntimeRegistry();

const buildGetDefaultSocketPathMainDepsHandler = createBuildGetDefaultSocketPathMainDepsHandler({
  platform: process.platform,
});
const getDefaultSocketPathMainDeps = buildGetDefaultSocketPathMainDepsHandler();
const getDefaultSocketPathHandler = createGetDefaultSocketPathHandler(getDefaultSocketPathMainDeps);

function getDefaultSocketPath(): string {
  return getDefaultSocketPathHandler();
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
let overlayModalInputExclusive = false;
let syncOverlayShortcutsForModal: (isActive: boolean) => void = () => {};

const handleModalInputStateChange = (isActive: boolean): void => {
  if (overlayModalInputExclusive === isActive) return;
  overlayModalInputExclusive = isActive;
  if (isActive) {
    const modalWindow = overlayManager.getModalWindow();
    if (modalWindow && !modalWindow.isDestroyed()) {
      modalWindow.setIgnoreMouseEvents(false);
      modalWindow.setAlwaysOnTop(true, 'screen-saver', 1);
      modalWindow.focus();
      if (!modalWindow.webContents.isFocused()) {
        modalWindow.webContents.focus();
      }
    }
  }
  syncOverlayShortcutsForModal(isActive);
};

const buildOverlayContentMeasurementStoreMainDepsHandler =
  createBuildOverlayContentMeasurementStoreMainDepsHandler({
    now: () => Date.now(),
    warn: (message: string) => logger.warn(message),
  });
const buildOverlayModalRuntimeMainDepsHandler = createBuildOverlayModalRuntimeMainDepsHandler({
  getMainWindow: () => overlayManager.getMainWindow(),
  getModalWindow: () => overlayManager.getModalWindow(),
  createModalWindow: () => createModalWindow(),
  getModalGeometry: () => getCurrentOverlayGeometry(),
  setModalWindowBounds: (geometry) => overlayManager.setModalWindowBounds(geometry),
});
const overlayContentMeasurementStoreMainDeps = buildOverlayContentMeasurementStoreMainDepsHandler();
const overlayContentMeasurementStore = createOverlayContentMeasurementStore(
  overlayContentMeasurementStoreMainDeps,
);
const overlayModalRuntime = createOverlayModalRuntimeService(
  buildOverlayModalRuntimeMainDepsHandler(),
  {
    onModalStateChange: (isActive: boolean) => handleModalInputStateChange(isActive),
  },
);
const appState = createAppState({
  mpvSocketPath: getDefaultSocketPath(),
  texthookerPort: DEFAULT_TEXTHOOKER_PORT,
});
const discordPresenceSessionStartedAtMs = Date.now();
let discordPresenceMediaDurationSec: number | null = null;

function refreshDiscordPresenceMediaDuration(): void {
  const client = appState.mpvClient;
  if (!client || !client.connected) return;
  void client
    .requestProperty('duration')
    .then((value) => {
      const numeric = Number(value);
      discordPresenceMediaDurationSec = Number.isFinite(numeric) && numeric > 0 ? numeric : null;
    })
    .catch(() => {
      discordPresenceMediaDurationSec = null;
    });
}

function publishDiscordPresence(): void {
  const discordPresenceService = appState.discordPresenceService;
  if (!discordPresenceService || getResolvedConfig().discordPresence.enabled !== true) {
    return;
  }

  refreshDiscordPresenceMediaDuration();
  discordPresenceService.publish({
    mediaTitle: appState.currentMediaTitle,
    mediaPath: appState.currentMediaPath,
    subtitleText: appState.currentSubText,
    currentTimeSec: appState.mpvClient?.currentTimePos ?? null,
    mediaDurationSec:
      discordPresenceMediaDurationSec ?? anilistMediaGuessRuntimeState.mediaDurationSec,
    paused: appState.playbackPaused,
    connected: Boolean(appState.mpvClient?.connected),
    sessionStartedAtMs: discordPresenceSessionStartedAtMs,
  });
}

function createDiscordRpcClient() {
  const discordRpc = require('discord-rpc') as {
    Client: new (opts: { transport: 'ipc' }) => {
      login: (opts: { clientId: string }) => Promise<void>;
      setActivity: (activity: Record<string, unknown>) => Promise<void>;
      clearActivity: () => Promise<void>;
      destroy: () => void;
    };
  };
  const client = new discordRpc.Client({ transport: 'ipc' });

  return {
    login: () => client.login({ clientId: DISCORD_PRESENCE_APP_ID }),
    setActivity: (activity: unknown) =>
      client.setActivity(activity as unknown as Record<string, unknown>),
    clearActivity: () => client.clearActivity(),
    destroy: () => client.destroy(),
  };
}

async function initializeDiscordPresenceService(): Promise<void> {
  if (getResolvedConfig().discordPresence.enabled !== true) {
    appState.discordPresenceService = null;
    return;
  }

  appState.discordPresenceService = createDiscordPresenceService({
    config: getResolvedConfig().discordPresence,
    createClient: () => createDiscordRpcClient(),
    logDebug: (message, meta) => logger.debug(message, meta),
  });
  await appState.discordPresenceService.start();
  publishDiscordPresence();
}
const ensureOverlayMpvSubtitlesHidden = createEnsureOverlayMpvSubtitlesHiddenHandler({
  getMpvClient: () => appState.mpvClient,
  getSavedSubVisibility: () => appState.overlaySavedMpvSubVisibility,
  setSavedSubVisibility: (visible) => {
    appState.overlaySavedMpvSubVisibility = visible;
  },
  getRevision: () => appState.overlayMpvSubVisibilityRevision,
  setRevision: (revision) => {
    appState.overlayMpvSubVisibilityRevision = revision;
  },
  setMpvSubVisibility: (visible) => {
    setMpvSubVisibilityRuntime(appState.mpvClient, visible);
  },
  logWarn: (message, error) => {
    logger.warn(message, error);
  },
});
const restoreOverlayMpvSubtitles = createRestoreOverlayMpvSubtitlesHandler({
  getSavedSubVisibility: () => appState.overlaySavedMpvSubVisibility,
  setSavedSubVisibility: (visible) => {
    appState.overlaySavedMpvSubVisibility = visible;
  },
  getRevision: () => appState.overlayMpvSubVisibilityRevision,
  setRevision: (revision) => {
    appState.overlayMpvSubVisibilityRevision = revision;
  },
  isMpvConnected: () => Boolean(appState.mpvClient?.connected),
  shouldKeepSuppressedFromVisibleOverlayBinding: () => shouldSuppressMpvSubtitlesForOverlay(),
  setMpvSubVisibility: (visible) => {
    setMpvSubVisibilityRuntime(appState.mpvClient, visible);
  },
});

function shouldSuppressMpvSubtitlesForOverlay(): boolean {
  return overlayManager.getVisibleOverlayVisible();
}

function syncOverlayMpvSubtitleSuppression(): void {
  if (shouldSuppressMpvSubtitlesForOverlay()) {
    void ensureOverlayMpvSubtitlesHidden();
    return;
  }

  restoreOverlayMpvSubtitles();
}

const buildImmersionMediaRuntimeMainDepsHandler = createBuildImmersionMediaRuntimeMainDepsHandler({
  getResolvedConfig: () => getResolvedConfig(),
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  getTracker: () => appState.immersionTracker,
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  logDebug: (message) => logger.debug(message),
  logInfo: (message) => logger.info(message),
});
const buildAnilistStateRuntimeMainDepsHandler = createBuildAnilistStateRuntimeMainDepsHandler({
  getClientSecretState: () => appState.anilistClientSecretState,
  setClientSecretState: (next) => {
    appState.anilistClientSecretState = transitionAnilistClientSecretState(
      appState.anilistClientSecretState,
      next,
    );
  },
  getRetryQueueState: () => appState.anilistRetryQueueState,
  setRetryQueueState: (next) => {
    appState.anilistRetryQueueState = transitionAnilistRetryQueueState(
      appState.anilistRetryQueueState,
      next,
    );
  },
  getUpdateQueueSnapshot: () => anilistUpdateQueue.getSnapshot(),
  clearStoredToken: () => anilistTokenStore.clearToken(),
  clearCachedAccessToken: () => {
    anilistCachedAccessToken = null;
  },
});
const buildConfigDerivedRuntimeMainDepsHandler = createBuildConfigDerivedRuntimeMainDepsHandler({
  getResolvedConfig: () => getResolvedConfig(),
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
  defaultJimakuLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  defaultJimakuMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
  defaultJimakuApiBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
});
const buildMainSubsyncRuntimeMainDepsHandler = createBuildMainSubsyncRuntimeMainDepsHandler({
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
const immersionMediaRuntime = createImmersionMediaRuntime(
  buildImmersionMediaRuntimeMainDepsHandler(),
);
const anilistStateRuntime = createAnilistStateRuntime(buildAnilistStateRuntimeMainDepsHandler());
const configDerivedRuntime = createConfigDerivedRuntime(buildConfigDerivedRuntimeMainDepsHandler());
const subsyncRuntime = createMainSubsyncRuntime(buildMainSubsyncRuntimeMainDepsHandler());
let autoPlayReadySignalMediaPath: string | null = null;
let autoPlayReadySignalGeneration = 0;

function maybeSignalPluginAutoplayReady(
  payload: SubtitleData,
  options?: { forceWhilePaused?: boolean },
): void {
  if (!payload.text.trim()) {
    return;
  }
  const mediaPath =
    appState.currentMediaPath?.trim() ||
    appState.mpvClient?.currentVideoPath?.trim() ||
    '__unknown__';
  const duplicateMediaSignal = autoPlayReadySignalMediaPath === mediaPath;
  const allowDuplicateWhilePaused =
    options?.forceWhilePaused === true && appState.playbackPaused !== false;
  if (duplicateMediaSignal && !allowDuplicateWhilePaused) {
    return;
  }
  autoPlayReadySignalMediaPath = mediaPath;
  const playbackGeneration = ++autoPlayReadySignalGeneration;
  const signalPluginAutoplayReady = (): void => {
    logger.debug(`[autoplay-ready] signaling mpv for media: ${mediaPath}`);
    sendMpvCommandRuntime(appState.mpvClient, ['script-message', 'subminer-autoplay-ready']);
  };
  signalPluginAutoplayReady();
  const isPlaybackPaused = async (client: {
    requestProperty: (property: string) => Promise<unknown>;
  }): Promise<boolean> => {
    try {
      const pauseProperty = await client.requestProperty('pause');
      if (typeof pauseProperty === 'boolean') {
        return pauseProperty;
      }
      if (typeof pauseProperty === 'string') {
        return pauseProperty.toLowerCase() !== 'no' && pauseProperty !== '0';
      }
      if (typeof pauseProperty === 'number') {
        return pauseProperty !== 0;
      }
      logger.debug(
        `[autoplay-ready] unrecognized pause property for media ${mediaPath}: ${String(pauseProperty)}`,
      );
    } catch (error) {
      logger.debug(
        `[autoplay-ready] failed to read pause property for media ${mediaPath}: ${(error as Error).message}`,
      );
    }
    return true;
  };

  // Fallback: repeatedly try to release pause for a short window in case startup
  // gate arming and tokenization-ready signal arrive out of order.
  const maxReleaseAttempts = options?.forceWhilePaused === true ? 14 : 3;
  const releaseRetryDelayMs = 200;
  const attemptRelease = (attempt: number): void => {
    void (async () => {
      if (
        autoPlayReadySignalMediaPath !== mediaPath ||
        playbackGeneration !== autoPlayReadySignalGeneration
      ) {
        return;
      }

      const mpvClient = appState.mpvClient;
      if (!mpvClient?.connected) {
        if (attempt < maxReleaseAttempts) {
          setTimeout(() => attemptRelease(attempt + 1), releaseRetryDelayMs);
        }
        return;
      }

      const shouldUnpause = await isPlaybackPaused(mpvClient);
      logger.debug(
        `[autoplay-ready] mpv paused before fallback attempt ${attempt} for ${mediaPath}: ${shouldUnpause}`,
      );
      if (!shouldUnpause) {
        if (attempt === 0) {
          logger.debug('[autoplay-ready] mpv already playing; no fallback unpause needed');
        }
        return;
      }

      signalPluginAutoplayReady();
      mpvClient.send({ command: ['set_property', 'pause', false] });
      if (attempt < maxReleaseAttempts) {
        setTimeout(() => attemptRelease(attempt + 1), releaseRetryDelayMs);
      }
    })();
  };
  attemptRelease(0);
}

let appTray: Tray | null = null;
const buildSubtitleProcessingControllerMainDepsHandler =
  createBuildSubtitleProcessingControllerMainDepsHandler({
    tokenizeSubtitle: async (text: string) => {
      return await tokenizeSubtitle(text);
    },
    emitSubtitle: (payload) => {
      appState.currentSubtitleData = payload;
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
const subtitleProcessingControllerMainDeps = buildSubtitleProcessingControllerMainDepsHandler();
const subtitleProcessingController = createSubtitleProcessingController(
  subtitleProcessingControllerMainDeps,
);
const overlayShortcutsRuntime = createOverlayShortcutsRuntimeService(
  createBuildOverlayShortcutsRuntimeMainDepsHandler({
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getShortcutsRegistered: () => appState.shortcutsRegistered,
    setShortcutsRegistered: (registered: boolean) => {
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
    copySubtitleMultiple: (timeoutMs: number) => {
      startPendingMultiCopy(timeoutMs);
    },
    copySubtitle: () => {
      copyCurrentSubtitle();
    },
    toggleSecondarySubMode: () => handleCycleSecondarySubMode(),
    updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
    triggerFieldGrouping: () => triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    mineSentenceCard: () => mineSentenceCard(),
    mineSentenceMultiple: (timeoutMs: number) => {
      startPendingMineSentenceMultiple(timeoutMs);
    },
    cancelPendingMultiCopy: () => {
      cancelPendingMultiCopy();
    },
    cancelPendingMineSentenceMultiple: () => {
      cancelPendingMineSentenceMultiple();
    },
  })(),
);
syncOverlayShortcutsForModal = (isActive: boolean): void => {
  if (isActive) {
    overlayShortcutsRuntime.unregisterOverlayShortcuts();
  } else {
    overlayShortcutsRuntime.syncOverlayShortcuts();
  }
};

const buildConfigHotReloadMessageMainDepsHandler = createBuildConfigHotReloadMessageMainDepsHandler(
  {
    showMpvOsd: (message) => showMpvOsd(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  },
);
const configHotReloadMessageMainDeps = buildConfigHotReloadMessageMainDepsHandler();
const notifyConfigHotReloadMessage = createConfigHotReloadMessageHandler(
  configHotReloadMessageMainDeps,
);
const buildWatchConfigPathMainDepsHandler = createBuildWatchConfigPathMainDepsHandler({
  fileExists: (targetPath) => fs.existsSync(targetPath),
  dirname: (targetPath) => path.dirname(targetPath),
  watchPath: (targetPath, listener) => fs.watch(targetPath, listener),
});
const watchConfigPathHandler = createWatchConfigPathHandler(buildWatchConfigPathMainDepsHandler());
const buildConfigHotReloadAppliedMainDepsHandler = createBuildConfigHotReloadAppliedMainDepsHandler(
  {
    setKeybindings: (keybindings) => {
      appState.keybindings = keybindings;
    },
    refreshGlobalAndOverlayShortcuts: () => {
      refreshGlobalAndOverlayShortcuts();
    },
    setSecondarySubMode: (mode) => {
      setSecondarySubMode(mode);
    },
    broadcastToOverlayWindows: (channel, payload) => {
      broadcastToOverlayWindows(channel, payload);
    },
    applyAnkiRuntimeConfigPatch: (patch) => {
      if (appState.ankiIntegration) {
        appState.ankiIntegration.applyRuntimeConfigPatch(patch);
      }
    },
  },
);
const buildConfigHotReloadRuntimeMainDepsHandler = createBuildConfigHotReloadRuntimeMainDepsHandler(
  {
    getCurrentConfig: () => getResolvedConfig(),
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    watchConfigPath: (configPath, onChange) => watchConfigPathHandler(configPath, onChange),
    setTimeout: (callback, delayMs) => setTimeout(callback, delayMs),
    clearTimeout: (timeout) => clearTimeout(timeout),
    debounceMs: 250,
    onHotReloadApplied: createConfigHotReloadAppliedHandler(
      buildConfigHotReloadAppliedMainDepsHandler(),
    ),
    onRestartRequired: (fields) =>
      notifyConfigHotReloadMessage(buildRestartRequiredConfigMessage(fields)),
    onInvalidConfig: notifyConfigHotReloadMessage,
    onValidationWarnings: (configPath, warnings) => {
      showDesktopNotification('SubMiner', {
        body: buildConfigWarningNotificationBody(configPath, warnings),
      });
      if (process.platform === 'darwin') {
        dialog.showErrorBox(
          'SubMiner config validation warning',
          buildConfigWarningDialogDetails(configPath, warnings),
        );
      }
    },
  },
);
const configHotReloadRuntime = createConfigHotReloadRuntime(
  buildConfigHotReloadRuntimeMainDepsHandler(),
);

const buildDictionaryRootsHandler = createBuildDictionaryRootsMainHandler({
  dirname: __dirname,
  appPath: app.getAppPath(),
  resourcesPath: process.resourcesPath,
  userDataPath: USER_DATA_PATH,
  appUserDataPath: app.getPath('userData'),
  homeDir: os.homedir(),
  cwd: process.cwd(),
  joinPath: (...parts) => path.join(...parts),
});
const buildFrequencyDictionaryRootsHandler = createBuildFrequencyDictionaryRootsMainHandler({
  dirname: __dirname,
  appPath: app.getAppPath(),
  resourcesPath: process.resourcesPath,
  userDataPath: USER_DATA_PATH,
  appUserDataPath: app.getPath('userData'),
  homeDir: os.homedir(),
  cwd: process.cwd(),
  joinPath: (...parts) => path.join(...parts),
});

const jlptDictionaryRuntime = createJlptDictionaryRuntimeService(
  createBuildJlptDictionaryRuntimeMainDepsHandler({
    isJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
    getDictionaryRoots: () => buildDictionaryRootsHandler(),
    getJlptDictionarySearchPaths,
    setJlptLevelLookup: (lookup) => {
      appState.jlptLevelLookup = lookup;
    },
    logInfo: (message) => logger.info(message),
  })(),
);

const frequencyDictionaryRuntime = createFrequencyDictionaryRuntimeService(
  createBuildFrequencyDictionaryRuntimeMainDepsHandler({
    isFrequencyDictionaryEnabled: () =>
      getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
    getDictionaryRoots: () => buildFrequencyDictionaryRootsHandler(),
    getFrequencyDictionarySearchPaths,
    getSourcePath: () => getResolvedConfig().subtitleStyle.frequencyDictionary.sourcePath,
    setFrequencyRankLookup: (lookup) => {
      appState.frequencyRankLookup = lookup;
    },
    logInfo: (message) => logger.info(message),
  })(),
);

const buildGetFieldGroupingResolverMainDepsHandler =
  createBuildGetFieldGroupingResolverMainDepsHandler({
    getResolver: () => appState.fieldGroupingResolver,
  });
const getFieldGroupingResolverMainDeps = buildGetFieldGroupingResolverMainDepsHandler();
const getFieldGroupingResolverHandler = createGetFieldGroupingResolverHandler(
  getFieldGroupingResolverMainDeps,
);

function getFieldGroupingResolver(): ((choice: KikuFieldGroupingChoice) => void) | null {
  return getFieldGroupingResolverHandler();
}

const buildSetFieldGroupingResolverMainDepsHandler =
  createBuildSetFieldGroupingResolverMainDepsHandler({
    setResolver: (resolver) => {
      appState.fieldGroupingResolver = resolver;
    },
    nextSequence: () => {
      appState.fieldGroupingResolverSequence += 1;
      return appState.fieldGroupingResolverSequence;
    },
    getSequence: () => appState.fieldGroupingResolverSequence,
  });
const setFieldGroupingResolverMainDeps = buildSetFieldGroupingResolverMainDepsHandler();
const setFieldGroupingResolverHandler = createSetFieldGroupingResolverHandler(
  setFieldGroupingResolverMainDeps,
);

function setFieldGroupingResolver(
  resolver: ((choice: KikuFieldGroupingChoice) => void) | null,
): void {
  setFieldGroupingResolverHandler(resolver);
}

const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntime<OverlayHostedModal>(
  createBuildFieldGroupingOverlayMainDepsHandler<OverlayHostedModal>({
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    getResolver: () => getFieldGroupingResolver(),
    setResolver: (resolver) => setFieldGroupingResolver(resolver),
    getRestoreVisibleOverlayOnModalClose: () =>
      overlayModalRuntime.getRestoreVisibleOverlayOnModalClose(),
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
      overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  })(),
);
const createFieldGroupingCallback = fieldGroupingOverlayRuntime.createFieldGroupingCallback;

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, 'subtitle-positions');

const mediaRuntime = createMediaRuntimeService(
  createBuildMediaRuntimeMainDepsHandler({
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
    broadcastToOverlayWindows: (channel, payload) => {
      broadcastToOverlayWindows(channel, payload);
    },
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    setCurrentMediaTitle: (title) => {
      appState.currentMediaTitle = title;
    },
  })(),
);

const overlayVisibilityRuntime = createOverlayVisibilityRuntimeService(
  createBuildOverlayVisibilityRuntimeMainDepsHandler({
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getWindowTracker: () => appState.windowTracker,
    getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      appState.trackerNotReadyWarningShown = shown;
    },
    updateVisibleOverlayBounds: (geometry: WindowGeometry) => updateVisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window) => {
      ensureOverlayWindowLevel(window);
    },
    syncPrimaryOverlayWindowLayer: (layer) => {
      syncPrimaryOverlayWindowLayer(layer);
    },
    enforceOverlayLayerOrder: () => {
      enforceOverlayLayerOrder();
    },
    syncOverlayShortcuts: () => {
      overlayShortcutsRuntime.syncOverlayShortcuts();
    },
    isMacOSPlatform: () => process.platform === 'darwin',
    showOverlayLoadingOsd: (message: string) => {
      showMpvOsd(message);
    },
    resolveFallbackBounds: () => {
      const cursorPoint = screen.getCursorScreenPoint();
      const display = screen.getDisplayNearestPoint(cursorPoint);
      const fallbackBounds = display.workArea;
      return {
        x: fallbackBounds.x,
        y: fallbackBounds.y,
        width: fallbackBounds.width,
        height: fallbackBounds.height,
      };
    },
  })(),
);

const buildGetRuntimeOptionsStateMainDepsHandler = createBuildGetRuntimeOptionsStateMainDepsHandler(
  {
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
  },
);
const getRuntimeOptionsStateMainDeps = buildGetRuntimeOptionsStateMainDepsHandler();
const getRuntimeOptionsStateHandler = createGetRuntimeOptionsStateHandler(
  getRuntimeOptionsStateMainDeps,
);

function getRuntimeOptionsState(): RuntimeOptionState[] {
  return getRuntimeOptionsStateHandler();
}

function getOverlayWindows(): BrowserWindow[] {
  return overlayManager.getOverlayWindows();
}

const buildRestorePreviousSecondarySubVisibilityMainDepsHandler =
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler({
    getMpvClient: () => appState.mpvClient,
  });
const restorePreviousSecondarySubVisibilityMainDeps =
  buildRestorePreviousSecondarySubVisibilityMainDepsHandler();
const restorePreviousSecondarySubVisibilityHandler =
  createRestorePreviousSecondarySubVisibilityHandler(restorePreviousSecondarySubVisibilityMainDeps);

function restorePreviousSecondarySubVisibility(): void {
  restorePreviousSecondarySubVisibilityHandler();
}

function broadcastToOverlayWindows(channel: string, ...args: unknown[]): void {
  overlayManager.broadcastToOverlayWindows(channel, ...args);
}

const buildBroadcastRuntimeOptionsChangedMainDepsHandler =
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler({
    broadcastRuntimeOptionsChangedRuntime,
    getRuntimeOptionsState: () => getRuntimeOptionsState(),
    broadcastToOverlayWindows: (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  });
const broadcastRuntimeOptionsChangedMainDeps = buildBroadcastRuntimeOptionsChangedMainDepsHandler();
const broadcastRuntimeOptionsChangedHandler = createBroadcastRuntimeOptionsChangedHandler(
  broadcastRuntimeOptionsChangedMainDeps,
);

function broadcastRuntimeOptionsChanged(): void {
  broadcastRuntimeOptionsChangedHandler();
}

const buildSendToActiveOverlayWindowMainDepsHandler =
  createBuildSendToActiveOverlayWindowMainDepsHandler({
    sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) =>
      overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });
const sendToActiveOverlayWindowMainDeps = buildSendToActiveOverlayWindowMainDepsHandler();
const sendToActiveOverlayWindowHandler = createSendToActiveOverlayWindowHandler(
  sendToActiveOverlayWindowMainDeps,
);

function sendToActiveOverlayWindow(
  channel: string,
  payload?: unknown,
  runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
): boolean {
  return sendToActiveOverlayWindowHandler(channel, payload, runtimeOptions);
}

const buildSetOverlayDebugVisualizationEnabledMainDepsHandler =
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler({
    setOverlayDebugVisualizationEnabledRuntime,
    getCurrentEnabled: () => appState.overlayDebugVisualizationEnabled,
    setCurrentEnabled: (next) => {
      appState.overlayDebugVisualizationEnabled = next;
    },
  });
const setOverlayDebugVisualizationEnabledMainDeps =
  buildSetOverlayDebugVisualizationEnabledMainDepsHandler();
const setOverlayDebugVisualizationEnabledHandler = createSetOverlayDebugVisualizationEnabledHandler(
  setOverlayDebugVisualizationEnabledMainDeps,
);

function setOverlayDebugVisualizationEnabled(enabled: boolean): void {
  setOverlayDebugVisualizationEnabledHandler(enabled);
}

const buildOpenRuntimeOptionsPaletteMainDepsHandler =
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler({
    openRuntimeOptionsPaletteRuntime: () => overlayModalRuntime.openRuntimeOptionsPalette(),
  });
const openRuntimeOptionsPaletteMainDeps = buildOpenRuntimeOptionsPaletteMainDepsHandler();
const openRuntimeOptionsPaletteHandler = createOpenRuntimeOptionsPaletteHandler(
  openRuntimeOptionsPaletteMainDeps,
);

function openRuntimeOptionsPalette(): void {
  openRuntimeOptionsPaletteHandler();
}

function getResolvedConfig() {
  return configService.getConfig();
}

function getRuntimeBooleanOption(
  id: 'subtitle.annotation.nPlusOne' | 'subtitle.annotation.jlpt' | 'subtitle.annotation.frequency',
  fallback: boolean,
): boolean {
  const value = appState.runtimeOptionsManager?.getOptionValue(id);
  return typeof value === 'boolean' ? value : fallback;
}

function shouldInitializeMecabForAnnotations(): boolean {
  const config = getResolvedConfig();
  const nPlusOneEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.nPlusOne',
    config.ankiConnect.nPlusOne.highlightEnabled,
  );
  const jlptEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.jlpt',
    config.subtitleStyle.enableJlpt,
  );
  const frequencyEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.frequency',
    config.subtitleStyle.frequencyDictionary.enabled,
  );
  return nPlusOneEnabled || jlptEnabled || frequencyEnabled;
}

const {
  getResolvedJellyfinConfig,
  getJellyfinClientInfo,
  reportJellyfinRemoteProgress,
  reportJellyfinRemoteStopped,
  handleJellyfinRemotePlay,
  handleJellyfinRemotePlaystate,
  handleJellyfinRemoteGeneralCommand,
  playJellyfinItemInMpv,
  startJellyfinRemoteSession,
  stopJellyfinRemoteSession,
  runJellyfinCommand,
  openJellyfinSetupWindow,
} = composeJellyfinRuntimeHandlers({
  getResolvedJellyfinConfigMainDeps: {
    getResolvedConfig: () => getResolvedConfig(),
    loadStoredSession: () => jellyfinTokenStore.loadSession(),
    getEnv: (name) => process.env[name],
  },
  getJellyfinClientInfoMainDeps: {
    getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
    getDefaultJellyfinConfig: () => DEFAULT_CONFIG.jellyfin,
  },
  waitForMpvConnectedMainDeps: {
    getMpvClient: () => appState.mpvClient,
    now: () => Date.now(),
    sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
  },
  launchMpvIdleForJellyfinPlaybackMainDeps: {
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
  },
  ensureMpvConnectedForJellyfinPlaybackMainDeps: {
    getMpvClient: () => appState.mpvClient,
    setMpvClient: (client) => {
      appState.mpvClient = client as MpvIpcClient | null;
    },
    createMpvClient: () => createMpvClientRuntimeService(),
    getAutoLaunchInFlight: () => jellyfinMpvAutoLaunchInFlight,
    setAutoLaunchInFlight: (promise) => {
      jellyfinMpvAutoLaunchInFlight = promise;
    },
    connectTimeoutMs: JELLYFIN_MPV_CONNECT_TIMEOUT_MS,
    autoLaunchTimeoutMs: JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS,
  },
  preloadJellyfinExternalSubtitlesMainDeps: {
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
    getMpvClient: () => appState.mpvClient,
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
    wait: (ms) => new Promise<void>((resolve) => setTimeout(resolve, ms)),
    logDebug: (message, error) => {
      logger.debug(message, error);
    },
  },
  playJellyfinItemInMpvMainDeps: {
    getMpvClient: () => appState.mpvClient,
    resolvePlaybackPlan: (params) =>
      resolveJellyfinPlaybackPlanRuntime(
        params.session,
        params.clientInfo,
        params.jellyfinConfig as ReturnType<typeof getResolvedJellyfinConfig>,
        {
          itemId: params.itemId,
          audioStreamIndex: params.audioStreamIndex ?? undefined,
          subtitleStreamIndex: params.subtitleStreamIndex ?? undefined,
        },
      ),
    applyJellyfinMpvDefaults: (mpvClient) => applyJellyfinMpvDefaults(mpvClient),
    sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
    armQuitOnDisconnect: () => {
      jellyfinPlayQuitOnDisconnectArmed = false;
      setTimeout(() => {
        jellyfinPlayQuitOnDisconnectArmed = true;
      }, 3000);
    },
    schedule: (callback, delayMs) => {
      setTimeout(callback, delayMs);
    },
    convertTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
    setActivePlayback: (state) => {
      activeJellyfinRemotePlayback = state as ActiveJellyfinRemotePlaybackState;
    },
    setLastProgressAtMs: (value) => {
      jellyfinRemoteLastProgressAtMs = value;
    },
    reportPlaying: (payload) => {
      void appState.jellyfinRemoteSession?.reportPlaying(payload);
    },
    showMpvOsd: (text) => {
      showMpvOsd(text);
    },
  },
  remoteComposerOptions: {
    getConfiguredSession: () => getConfiguredJellyfinSession(getResolvedJellyfinConfig()),
    logWarn: (message) => logger.warn(message),
    getMpvClient: () => appState.mpvClient,
    sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
    jellyfinTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
    getActivePlayback: () => activeJellyfinRemotePlayback,
    clearActivePlayback: () => {
      activeJellyfinRemotePlayback = null;
    },
    getSession: () => appState.jellyfinRemoteSession,
    getNow: () => Date.now(),
    getLastProgressAtMs: () => jellyfinRemoteLastProgressAtMs,
    setLastProgressAtMs: (value) => {
      jellyfinRemoteLastProgressAtMs = value;
    },
    progressIntervalMs: JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS,
    ticksPerSecond: JELLYFIN_TICKS_PER_SECOND,
    logDebug: (message, error) => logger.debug(message, error),
  },
  handleJellyfinAuthCommandsMainDeps: {
    patchRawConfig: (patch) => {
      configService.patchRawConfig(patch);
    },
    authenticateWithPassword: (serverUrl, username, password, clientInfo) =>
      authenticateWithPasswordRuntime(serverUrl, username, password, clientInfo),
    saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
    clearStoredSession: () => jellyfinTokenStore.clearSession(),
    logInfo: (message) => logger.info(message),
  },
  handleJellyfinListCommandsMainDeps: {
    listJellyfinLibraries: (session, clientInfo) =>
      listJellyfinLibrariesRuntime(session, clientInfo),
    listJellyfinItems: (session, clientInfo, params) =>
      listJellyfinItemsRuntime(session, clientInfo, params),
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
    writeJellyfinPreviewAuth: (responsePath, payload) => {
      fs.mkdirSync(path.dirname(responsePath), { recursive: true });
      fs.writeFileSync(responsePath, JSON.stringify(payload, null, 2), 'utf-8');
    },
    logInfo: (message) => logger.info(message),
  },
  handleJellyfinPlayCommandMainDeps: {
    logWarn: (message) => logger.warn(message),
  },
  handleJellyfinRemoteAnnounceCommandMainDeps: {
    getRemoteSession: () => appState.jellyfinRemoteSession,
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
  },
  startJellyfinRemoteSessionMainDeps: {
    getCurrentSession: () => appState.jellyfinRemoteSession,
    setCurrentSession: (session) => {
      appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
    },
    createRemoteSessionService: (options) => new JellyfinRemoteSessionService(options),
    defaultDeviceId: DEFAULT_CONFIG.jellyfin.deviceId,
    defaultClientName: DEFAULT_CONFIG.jellyfin.clientName,
    defaultClientVersion: DEFAULT_CONFIG.jellyfin.clientVersion,
    logInfo: (message) => logger.info(message),
    logWarn: (message, details) => logger.warn(message, details),
  },
  stopJellyfinRemoteSessionMainDeps: {
    getCurrentSession: () => appState.jellyfinRemoteSession,
    setCurrentSession: (session) => {
      appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
    },
    clearActivePlayback: () => {
      activeJellyfinRemotePlayback = null;
    },
  },
  runJellyfinCommandMainDeps: {
    defaultServerUrl: DEFAULT_CONFIG.jellyfin.serverUrl,
  },
  maybeFocusExistingJellyfinSetupWindowMainDeps: {
    getSetupWindow: () => appState.jellyfinSetupWindow,
  },
  openJellyfinSetupWindowMainDeps: {
    createSetupWindow: () =>
      new BrowserWindow({
        width: 520,
        height: 560,
        title: 'Jellyfin Setup',
        show: true,
        autoHideMenuBar: true,
        webPreferences: {
          nodeIntegration: false,
          contextIsolation: true,
        },
      }),
    buildSetupFormHtml: (defaultServer, defaultUser) =>
      buildJellyfinSetupFormHtml(defaultServer, defaultUser),
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: (server, username, password, clientInfo) =>
      authenticateWithPasswordRuntime(server, username, password, clientInfo),
    saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
    patchJellyfinConfig: (session) => {
      configService.patchRawConfig({
        jellyfin: {
          enabled: true,
          serverUrl: session.serverUrl,
          username: session.username,
        },
      });
    },
    logInfo: (message) => logger.info(message),
    logError: (message, error) => logger.error(message, error),
    showMpvOsd: (message) => showMpvOsd(message),
    clearSetupWindow: () => {
      appState.jellyfinSetupWindow = null;
    },
    setSetupWindow: (window) => {
      appState.jellyfinSetupWindow = window as BrowserWindow;
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
  },
});

const {
  notifyAnilistSetup,
  consumeAnilistSetupTokenFromUrl,
  handleAnilistSetupProtocolUrl,
  registerSubminerProtocolClient,
} = composeAnilistSetupHandlers({
  notifyDeps: {
    hasMpvClient: () => Boolean(appState.mpvClient),
    showMpvOsd: (message) => showMpvOsd(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
    logInfo: (message) => logger.info(message),
  },
  consumeTokenDeps: {
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
  },
  handleProtocolDeps: {
    consumeAnilistSetupTokenFromUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
    logWarn: (message, details) => logger.warn(message, details),
  },
  registerProtocolClientDeps: {
    isDefaultApp: () => Boolean(process.defaultApp),
    getArgv: () => process.argv,
    execPath: process.execPath,
    resolvePath: (value) => path.resolve(value),
    setAsDefaultProtocolClient: (scheme, appPath, args) =>
      appPath
        ? app.setAsDefaultProtocolClient(scheme, appPath, args)
        : app.setAsDefaultProtocolClient(scheme),
    logWarn: (message, details) => logger.warn(message, details),
  },
});

const maybeFocusExistingAnilistSetupWindow = createMaybeFocusExistingAnilistSetupWindowHandler({
  getSetupWindow: () => appState.anilistSetupWindow,
});
const buildOpenAnilistSetupWindowMainDepsHandler = createBuildOpenAnilistSetupWindowMainDepsHandler(
  {
    maybeFocusExistingSetupWindow: maybeFocusExistingAnilistSetupWindow,
    createSetupWindow: () =>
      new BrowserWindow({
        width: 1000,
        height: 760,
        title: 'Anilist Setup',
        show: true,
        autoHideMenuBar: true,
        webPreferences: {
          nodeIntegration: false,
          contextIsolation: true,
        },
      }),
    buildAuthorizeUrl: () =>
      buildAnilistSetupUrl({
        authorizeUrl: ANILIST_SETUP_CLIENT_ID_URL,
        clientId: ANILIST_DEFAULT_CLIENT_ID,
        responseType: ANILIST_SETUP_RESPONSE_TYPE,
      }),
    consumeCallbackUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
    openSetupInBrowser: (authorizeUrl) =>
      openAnilistSetupInBrowser({
        authorizeUrl,
        openExternal: (url) => shell.openExternal(url),
        logError: (message, error) => logger.error(message, error),
      }),
    loadManualTokenEntry: (setupWindow, authorizeUrl) =>
      loadAnilistManualTokenEntry({
        setupWindow: setupWindow as BrowserWindow,
        authorizeUrl,
        developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
        logWarn: (message, data) => logger.warn(message, data),
      }),
    redirectUri: ANILIST_REDIRECT_URI,
    developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
    isAllowedExternalUrl: (url) => isAllowedAnilistExternalUrl(url),
    isAllowedNavigationUrl: (url) => isAllowedAnilistSetupNavigationUrl(url),
    logWarn: (message, details) => logger.warn(message, details),
    logError: (message, details) => logger.error(message, details),
    clearSetupWindow: () => {
      appState.anilistSetupWindow = null;
    },
    setSetupPageOpened: (opened) => {
      appState.anilistSetupPageOpened = opened;
    },
    setSetupWindow: (setupWindow) => {
      appState.anilistSetupWindow = setupWindow as BrowserWindow;
    },
    openExternal: (url) => {
      void shell.openExternal(url);
    },
  },
);

function openAnilistSetupWindow(): void {
  createOpenAnilistSetupWindowHandler(buildOpenAnilistSetupWindowMainDepsHandler())();
}

const {
  refreshAnilistClientSecretState,
  getCurrentAnilistMediaKey,
  resetAnilistMediaTracking,
  getAnilistMediaGuessRuntimeState,
  setAnilistMediaGuessRuntimeState,
  resetAnilistMediaGuessState,
  maybeProbeAnilistDuration,
  ensureAnilistMediaGuess,
  processNextAnilistRetryUpdate,
  maybeRunAnilistPostWatchUpdate,
} = composeAnilistTrackingHandlers({
  refreshClientSecretMainDeps: {
    getResolvedConfig: () => getResolvedConfig(),
    isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
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
  },
  getCurrentMediaKeyMainDeps: {
    getCurrentMediaPath: () => appState.currentMediaPath,
  },
  resetMediaTrackingMainDeps: {
    setMediaKey: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaKey: value },
      );
    },
    setMediaDurationSec: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaDurationSec: value },
      );
    },
    setMediaGuess: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuess: value },
      );
    },
    setMediaGuessPromise: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuessPromise: value },
      );
    },
    setLastDurationProbeAtMs: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { lastDurationProbeAtMs: value },
      );
    },
  },
  getMediaGuessRuntimeStateMainDeps: {
    getMediaKey: () => anilistMediaGuessRuntimeState.mediaKey,
    getMediaDurationSec: () => anilistMediaGuessRuntimeState.mediaDurationSec,
    getMediaGuess: () => anilistMediaGuessRuntimeState.mediaGuess,
    getMediaGuessPromise: () => anilistMediaGuessRuntimeState.mediaGuessPromise,
    getLastDurationProbeAtMs: () => anilistMediaGuessRuntimeState.lastDurationProbeAtMs,
  },
  setMediaGuessRuntimeStateMainDeps: {
    setMediaKey: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaKey: value },
      );
    },
    setMediaDurationSec: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaDurationSec: value },
      );
    },
    setMediaGuess: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuess: value },
      );
    },
    setMediaGuessPromise: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuessPromise: value },
      );
    },
    setLastDurationProbeAtMs: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { lastDurationProbeAtMs: value },
      );
    },
  },
  resetMediaGuessStateMainDeps: {
    setMediaGuess: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuess: value },
      );
    },
    setMediaGuessPromise: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuessPromise: value },
      );
    },
  },
  maybeProbeDurationMainDeps: {
    getState: () => getAnilistMediaGuessRuntimeState(),
    setState: (state) => {
      setAnilistMediaGuessRuntimeState(state);
    },
    durationRetryIntervalMs: ANILIST_DURATION_RETRY_INTERVAL_MS,
    now: () => Date.now(),
    requestMpvDuration: async () => appState.mpvClient?.requestProperty('duration'),
    logWarn: (message, error) => logger.warn(message, error),
  },
  ensureMediaGuessMainDeps: {
    getState: () => getAnilistMediaGuessRuntimeState(),
    setState: (state) => {
      setAnilistMediaGuessRuntimeState(state);
    },
    resolveMediaPathForJimaku: (currentMediaPath) =>
      mediaRuntime.resolveMediaPathForJimaku(currentMediaPath),
    getCurrentMediaPath: () => appState.currentMediaPath,
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
  },
  processNextRetryUpdateMainDeps: {
    nextReady: () => anilistUpdateQueue.nextReady(),
    refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
    setLastAttemptAt: (value) => {
      appState.anilistRetryQueueState = transitionAnilistRetryQueueLastAttemptAt(
        appState.anilistRetryQueueState,
        value,
      );
    },
    setLastError: (value) => {
      appState.anilistRetryQueueState = transitionAnilistRetryQueueLastError(
        appState.anilistRetryQueueState,
        value,
      );
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
  },
  maybeRunPostWatchUpdateMainDeps: {
    getInFlight: () => anilistUpdateInFlightState.inFlight,
    setInFlight: (value) => {
      anilistUpdateInFlightState = transitionAnilistUpdateInFlightState(
        anilistUpdateInFlightState,
        value,
      );
    },
    getResolvedConfig: () => getResolvedConfig(),
    isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
    getCurrentMediaKey: () => getCurrentAnilistMediaKey(),
    hasMpvClient: () => Boolean(appState.mpvClient),
    getTrackedMediaKey: () => anilistMediaGuessRuntimeState.mediaKey,
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
  },
});

function refreshAnilistClientSecretStateIfEnabled(options?: {
  force?: boolean;
}): Promise<string | null> {
  if (!isAnilistTrackingEnabled(getResolvedConfig())) {
    return Promise.resolve(null);
  }
  return refreshAnilistClientSecretState(options);
}

const rememberAnilistAttemptedUpdate = (key: string): void => {
  rememberAnilistAttemptedUpdateKey(
    anilistAttemptedUpdateKeys,
    key,
    ANILIST_MAX_ATTEMPTED_UPDATE_KEYS,
  );
};

const buildLoadSubtitlePositionMainDepsHandler = createBuildLoadSubtitlePositionMainDepsHandler({
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
const loadSubtitlePositionMainDeps = buildLoadSubtitlePositionMainDepsHandler();
const loadSubtitlePosition = createLoadSubtitlePositionHandler(loadSubtitlePositionMainDeps);

const buildSaveSubtitlePositionMainDepsHandler = createBuildSaveSubtitlePositionMainDepsHandler({
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
const saveSubtitlePositionMainDeps = buildSaveSubtitlePositionMainDepsHandler();
const saveSubtitlePosition = createSaveSubtitlePositionHandler(saveSubtitlePositionMainDeps);

registerSubminerProtocolClient();
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
  },
  onWillQuitCleanupMainDeps: {
    destroyTray: () => destroyTray(),
    stopConfigHotReload: () => configHotReloadRuntime.stop(),
    restorePreviousSecondarySubVisibility: () => restorePreviousSecondarySubVisibility(),
    restoreMpvSubVisibility: () => {
      restoreOverlayMpvSubtitles();
    },
    unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
    stopSubtitleWebsocket: () => subtitleWsService.stop(),
    stopTexthookerService: () => texthookerService.stop(),
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    clearYomitanParserState: () => {
      appState.yomitanParserWindow = null;
      appState.yomitanParserReadyPromise = null;
      appState.yomitanParserInitPromise = null;
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
    stopJellyfinRemoteSession: () => stopJellyfinRemoteSession(),
    stopDiscordPresenceService: () => {
      void appState.discordPresenceService?.stop();
      appState.discordPresenceService = null;
    },
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

const immersionTrackerStartupMainDeps: Parameters<
  typeof createBuildImmersionTrackerStartupMainDepsHandler
>[0] = {
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
};
const createImmersionTrackerStartup = createImmersionTrackerStartupHandler(
  createBuildImmersionTrackerStartupMainDepsHandler(immersionTrackerStartupMainDeps)(),
);
let hasAttemptedImmersionTrackerStartup = false;
const ensureImmersionTrackerStarted = (): void => {
  if (hasAttemptedImmersionTrackerStartup || appState.immersionTracker) {
    return;
  }
  hasAttemptedImmersionTrackerStartup = true;
  createImmersionTrackerStartup();
};

const { reloadConfig: reloadConfigHandler, appReadyRuntimeRunner } = composeAppReadyRuntime({
  reloadConfigMainDeps: {
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    logInfo: (message) => appLogger.logInfo(message),
    logWarning: (message) => appLogger.logWarning(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
    startConfigHotReload: () => configHotReloadRuntime.start(),
    refreshAnilistClientSecretState: (options) => refreshAnilistClientSecretStateIfEnabled(options),
    failHandlers: {
      logError: (details) => logger.error(details),
      showErrorBox: (title, details) => dialog.showErrorBox(title, details),
      quit: () => app.quit(),
    },
  },
  criticalConfigErrorMainDeps: {
    getConfigPath: () => configService.getConfigPath(),
    failHandlers: {
      logError: (message) => logger.error(message),
      showErrorBox: (title, message) => dialog.showErrorBox(title, message),
      quit: () => app.quit(),
    },
  },
  appReadyRuntimeMainDeps: {
    loadSubtitlePosition: () => loadSubtitlePosition(),
    resolveKeybindings: () => {
      appState.keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);
    },
    createMpvClient: () => {
      appState.mpvClient = createMpvClientRuntimeService();
    },
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
          getSubtitleStyleConfig: () => configService.getConfig().subtitleStyle,
          onOptionsChanged: () => {
            subtitleProcessingController.invalidateTokenizationCache();
            broadcastRuntimeOptionsChanged();
            refreshOverlayShortcuts();
          },
        },
      );
    },
    setSecondarySubMode: (mode: SecondarySubMode) => {
      setSecondarySubMode(mode);
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
    shouldSkipHeavyStartup: () =>
      Boolean(appState.initialArgs && shouldRunSettingsOnlyStartup(appState.initialArgs)),
    createImmersionTracker: () => {
      ensureImmersionTrackerStarted();
    },
    logDebug: (message: string) => {
      logger.debug(message);
    },
    now: () => Date.now(),
  },
  immersionTrackerStartupMainDeps,
});

const { appLifecycleRuntimeRunner, runAndApplyStartupState } =
  runtimeRegistry.startup.createStartupRuntimeHandlers<
    CliArgs,
    StartupState,
    ReturnType<typeof createStartupBootstrapRuntimeDeps>
  >({
    appLifecycleRuntimeRunnerMainDeps: {
      app,
      platform: process.platform,
      shouldStartApp: (nextArgs: CliArgs) => shouldStartApp(nextArgs),
      parseArgs: (argv: string[]) => parseArgs(argv),
      handleCliCommand: (nextArgs: CliArgs, source: CliCommandSource) =>
        handleCliCommand(nextArgs, source),
      printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
      logNoRunningInstance: () => appLogger.logNoRunningInstance(),
      onReady: appReadyRuntimeRunner,
      onWillQuitCleanup: () => onWillQuitCleanupHandler(),
      shouldRestoreWindowsOnActivate: () => shouldRestoreWindowsOnActivateHandler(),
      restoreWindowsOnActivate: () => restoreWindowsOnActivateHandler(),
      shouldQuitOnWindowAllClosed: () => !appState.backgroundMode,
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
      setExitCode: (code) => {
        process.exitCode = code;
      },
      quitApp: () => app.quit(),
      logGenerateConfigError: (message) => logger.error(message),
      startAppLifecycle,
    }),
    createStartupBootstrapRuntimeDeps: (deps) => createStartupBootstrapRuntimeDeps(deps),
    runStartupBootstrapRuntime,
    applyStartupState: (startupState) => applyStartupState(appState, startupState),
  });

runAndApplyStartupState();
if (isAnilistTrackingEnabled(getResolvedConfig())) {
  void refreshAnilistClientSecretStateIfEnabled({ force: true });
  anilistStateRuntime.refreshRetryQueueState();
}
void initializeDiscordPresenceService();

const handleCliCommand = createCliCommandRuntimeHandler({
  handleTexthookerOnlyModeTransitionMainDeps: {
    isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
    setTexthookerOnlyMode: (enabled) => {
      appState.texthookerOnlyMode = enabled;
    },
    commandNeedsOverlayRuntime: (inputArgs) => commandNeedsOverlayRuntime(inputArgs),
    startBackgroundWarmups: () => startBackgroundWarmups(),
    logInfo: (message: string) => logger.info(message),
  },
  createCliCommandContext: () => createCliCommandContextHandler(),
  handleCliCommandRuntimeServiceWithContext: (args, source, cliContext) =>
    handleCliCommandRuntimeServiceWithContext(args, source, cliContext),
});

const handleInitialArgsRuntimeHandler = createInitialArgsRuntimeHandler({
  getInitialArgs: () => appState.initialArgs,
  isBackgroundMode: () => appState.backgroundMode,
  ensureTray: () => ensureTray(),
  isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
  hasImmersionTracker: () => Boolean(appState.immersionTracker),
  getMpvClient: () => appState.mpvClient,
  logInfo: (message) => logger.info(message),
  handleCliCommand: (args, source) => handleCliCommand(args, source),
});

function handleInitialArgs(): void {
  handleInitialArgsRuntimeHandler();
}

const {
  bindMpvClientEventHandlers,
  createMpvClientRuntimeService: createMpvClientRuntimeServiceHandler,
  updateMpvSubtitleRenderMetrics: updateMpvSubtitleRenderMetricsHandler,
  tokenizeSubtitle,
  createMecabTokenizerAndCheck,
  prewarmSubtitleDictionaries,
  startBackgroundWarmups,
} = composeMpvRuntimeHandlers<
  MpvIpcClient,
  ReturnType<typeof createTokenizerDepsRuntime>,
  SubtitleData
>({
  bindMpvMainEventHandlersMainDeps: {
    appState,
    getQuitOnDisconnectArmed: () => jellyfinPlayQuitOnDisconnectArmed,
    scheduleQuitCheck: (callback) => {
      setTimeout(callback, 500);
    },
    quitApp: () => app.quit(),
    reportJellyfinRemoteStopped: () => {
      void reportJellyfinRemoteStopped();
    },
    maybeRunAnilistPostWatchUpdate: () => maybeRunAnilistPostWatchUpdate(),
    logSubtitleTimingError: (message, error) => logger.error(message, error),
    broadcastToOverlayWindows: (channel, payload) => {
      broadcastToOverlayWindows(channel, payload);
    },
    onSubtitleChange: (text) => {
      subtitleProcessingController.onSubtitleChange(text);
    },
    refreshDiscordPresence: () => {
      publishDiscordPresence();
    },
    ensureImmersionTrackerInitialized: () => {
      ensureImmersionTrackerStarted();
    },
    updateCurrentMediaPath: (path) => {
      autoPlayReadySignalMediaPath = null;
      if (path) {
        ensureImmersionTrackerStarted();
      }
      mediaRuntime.updateCurrentMediaPath(path);
    },
    restoreMpvSubVisibility: () => {
      restoreOverlayMpvSubtitles();
    },
    getCurrentAnilistMediaKey: () => getCurrentAnilistMediaKey(),
    resetAnilistMediaTracking: (mediaKey) => {
      resetAnilistMediaTracking(mediaKey);
    },
    maybeProbeAnilistDuration: (mediaKey) => {
      void maybeProbeAnilistDuration(mediaKey);
    },
    ensureAnilistMediaGuess: (mediaKey) => {
      void ensureAnilistMediaGuess(mediaKey);
    },
    syncImmersionMediaState: () => {
      immersionMediaRuntime.syncFromCurrentMediaState();
    },
    updateCurrentMediaTitle: (title) => {
      mediaRuntime.updateCurrentMediaTitle(title);
    },
    resetAnilistMediaGuessState: () => {
      resetAnilistMediaGuessState();
    },
    reportJellyfinRemoteProgress: (forceImmediate) => {
      void reportJellyfinRemoteProgress(forceImmediate);
    },
    updateSubtitleRenderMetrics: (patch) => {
      updateMpvSubtitleRenderMetrics(patch as Partial<MpvSubtitleRenderMetrics>);
    },
    syncOverlayMpvSubtitleSuppression: () => {
      syncOverlayMpvSubtitleSuppression();
    },
  },
  mpvClientRuntimeServiceFactoryMainDeps: {
    createClient: MpvIpcClient,
    getSocketPath: () => appState.mpvSocketPath,
    getResolvedConfig: () => getResolvedConfig(),
    isAutoStartOverlayEnabled: () => appState.autoStartOverlay,
    setOverlayVisible: (visible: boolean) => setOverlayVisible(visible),
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getReconnectTimer: () => appState.reconnectTimer,
    setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
      appState.reconnectTimer = timer;
    },
  },
  updateMpvSubtitleRenderMetricsMainDeps: {
    getCurrentMetrics: () => appState.mpvSubtitleRenderMetrics,
    setCurrentMetrics: (metrics) => {
      appState.mpvSubtitleRenderMetrics = metrics;
    },
    applyPatch: (current, patch) => applyMpvSubtitleRenderMetricsPatch(current, patch),
    broadcastMetrics: () => {
      // no renderer consumer for subtitle render metrics updates at present
    },
  },
  tokenizer: {
    buildTokenizerDepsMainDeps: {
      getYomitanExt: () => appState.yomitanExt,
      getYomitanParserWindow: () => appState.yomitanParserWindow,
      setYomitanParserWindow: (window) => {
        appState.yomitanParserWindow = window as BrowserWindow | null;
      },
      getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
      setYomitanParserReadyPromise: (promise) => {
        appState.yomitanParserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
      setYomitanParserInitPromise: (promise) => {
        appState.yomitanParserInitPromise = promise;
      },
      isKnownWord: (text) => Boolean(appState.ankiIntegration?.isKnownWord(text)),
      recordLookup: (hit) => {
        ensureImmersionTrackerStarted();
        appState.immersionTracker?.recordLookup(hit);
      },
      getKnownWordMatchMode: () =>
        appState.ankiIntegration?.getKnownWordMatchMode() ??
        getResolvedConfig().ankiConnect.nPlusOne.matchMode,
      getNPlusOneEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.nPlusOne',
          getResolvedConfig().ankiConnect.nPlusOne.highlightEnabled,
        ),
      getMinSentenceWordsForNPlusOne: () =>
        getResolvedConfig().ankiConnect.nPlusOne.minSentenceWords,
      getJlptLevel: (text) => appState.jlptLevelLookup(text),
      getJlptEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.jlpt',
          getResolvedConfig().subtitleStyle.enableJlpt,
        ),
      getFrequencyDictionaryEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.frequency',
          getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
        ),
      getFrequencyDictionaryMatchMode: () =>
        getResolvedConfig().subtitleStyle.frequencyDictionary.matchMode,
      getFrequencyRank: (text) => appState.frequencyRankLookup(text),
      getYomitanGroupDebugEnabled: () => appState.overlayDebugVisualizationEnabled,
      getMecabTokenizer: () => appState.mecabTokenizer,
      onTokenizationReady: (text) => {
        maybeSignalPluginAutoplayReady({ text, tokens: null }, { forceWhilePaused: true });
      },
    },
    createTokenizerRuntimeDeps: (deps) =>
      createTokenizerDepsRuntime(deps as Parameters<typeof createTokenizerDepsRuntime>[0]),
    tokenizeSubtitle: (text, deps) => tokenizeSubtitleCore(text, deps),
    createMecabTokenizerAndCheckMainDeps: {
      getMecabTokenizer: () => appState.mecabTokenizer,
      setMecabTokenizer: (tokenizer) => {
        appState.mecabTokenizer = tokenizer as MecabTokenizer | null;
      },
      createMecabTokenizer: () => new MecabTokenizer(),
      checkAvailability: async (tokenizer) => (tokenizer as MecabTokenizer).checkAvailability(),
    },
    prewarmSubtitleDictionariesMainDeps: {
      ensureJlptDictionaryLookup: () => jlptDictionaryRuntime.ensureJlptDictionaryLookup(),
      ensureFrequencyDictionaryLookup: () =>
        frequencyDictionaryRuntime.ensureFrequencyDictionaryLookup(),
      showMpvOsd: (message: string) => showMpvOsd(message),
      shouldShowOsdNotification: () => {
        const type = getResolvedConfig().ankiConnect.behavior.notificationType;
        return type === 'osd' || type === 'both';
      },
    },
  },
  warmups: {
    launchBackgroundWarmupTaskMainDeps: {
      now: () => Date.now(),
      logDebug: (message) => logger.debug(message),
      logWarn: (message) => logger.warn(message),
    },
    startBackgroundWarmupsMainDeps: {
      getStarted: () => backgroundWarmupsStarted,
      setStarted: (started) => {
        backgroundWarmupsStarted = started;
      },
      isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
      ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded().then(() => {}),
      shouldWarmupMecab: () => {
        const startupWarmups = getResolvedConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        if (!startupWarmups.mecab) {
          return false;
        }
        return shouldInitializeMecabForAnnotations();
      },
      shouldWarmupYomitanExtension: () => getResolvedConfig().startupWarmups.yomitanExtension,
      shouldWarmupSubtitleDictionaries: () => {
        const startupWarmups = getResolvedConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        return startupWarmups.subtitleDictionaries;
      },
      shouldWarmupJellyfinRemoteSession: () => {
        const startupWarmups = getResolvedConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        return startupWarmups.jellyfinRemoteSession;
      },
      shouldAutoConnectJellyfinRemote: () => {
        const jellyfin = getResolvedConfig().jellyfin;
        return (
          jellyfin.enabled && jellyfin.remoteControlEnabled && jellyfin.remoteControlAutoConnect
        );
      },
      startJellyfinRemoteSession: () => startJellyfinRemoteSession(),
      logDebug: (message) => logger.debug(message),
    },
  },
});

function createMpvClientRuntimeService(): MpvIpcClient {
  return createMpvClientRuntimeServiceHandler() as MpvIpcClient;
}

function updateMpvSubtitleRenderMetrics(patch: Partial<MpvSubtitleRenderMetrics>): void {
  updateMpvSubtitleRenderMetricsHandler(patch);
}

let lastOverlayWindowGeometry: WindowGeometry | null = null;

function getOverlayGeometryFallback(): WindowGeometry {
  const cursorPoint = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(cursorPoint);
  const bounds = display.workArea;
  return {
    x: bounds.x,
    y: bounds.y,
    width: bounds.width,
    height: bounds.height,
  };
}

function getCurrentOverlayGeometry(): WindowGeometry {
  if (lastOverlayWindowGeometry) return lastOverlayWindowGeometry;
  const trackerGeometry = appState.windowTracker?.getGeometry();
  if (trackerGeometry) return trackerGeometry;
  return getOverlayGeometryFallback();
}

function applyOverlayRegions(geometry: WindowGeometry): void {
  lastOverlayWindowGeometry = geometry;
  overlayManager.setOverlayWindowBounds(geometry);
  overlayManager.setModalWindowBounds(geometry);
}

const buildUpdateVisibleOverlayBoundsMainDepsHandler =
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler({
    setOverlayWindowBounds: (geometry) => applyOverlayRegions(geometry),
  });
const updateVisibleOverlayBoundsMainDeps = buildUpdateVisibleOverlayBoundsMainDepsHandler();
const updateVisibleOverlayBounds = createUpdateVisibleOverlayBoundsHandler(
  updateVisibleOverlayBoundsMainDeps,
);

const buildEnsureOverlayWindowLevelMainDepsHandler =
  createBuildEnsureOverlayWindowLevelMainDepsHandler({
    ensureOverlayWindowLevelCore: (window) => ensureOverlayWindowLevelCore(window as BrowserWindow),
  });
const ensureOverlayWindowLevelMainDeps = buildEnsureOverlayWindowLevelMainDepsHandler();
const ensureOverlayWindowLevel = createEnsureOverlayWindowLevelHandler(
  ensureOverlayWindowLevelMainDeps,
);

function syncPrimaryOverlayWindowLayer(layer: 'visible'): void {
  const mainWindow = overlayManager.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed()) return;
  syncOverlayWindowLayer(mainWindow, layer);
}

const buildEnforceOverlayLayerOrderMainDepsHandler =
  createBuildEnforceOverlayLayerOrderMainDepsHandler({
    enforceOverlayLayerOrderCore: (params) =>
      enforceOverlayLayerOrderCore({
        visibleOverlayVisible: params.visibleOverlayVisible,
        mainWindow: params.mainWindow as BrowserWindow | null,
        ensureOverlayWindowLevel: (window) =>
          params.ensureOverlayWindowLevel(window as BrowserWindow),
      }),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getMainWindow: () => overlayManager.getMainWindow(),
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window as BrowserWindow),
  });
const enforceOverlayLayerOrderMainDeps = buildEnforceOverlayLayerOrderMainDepsHandler();
const enforceOverlayLayerOrder = createEnforceOverlayLayerOrderHandler(
  enforceOverlayLayerOrderMainDeps,
);

async function loadYomitanExtension(): Promise<Extension | null> {
  const extension = await yomitanExtensionRuntime.loadYomitanExtension();
  if (extension) {
    await syncYomitanDefaultProfileAnkiServer();
  }
  return extension;
}

async function ensureYomitanExtensionLoaded(): Promise<Extension | null> {
  const extension = await yomitanExtensionRuntime.ensureYomitanExtensionLoaded();
  if (extension) {
    await syncYomitanDefaultProfileAnkiServer();
  }
  return extension;
}

let lastSyncedYomitanAnkiServer: string | null = null;

function getPreferredYomitanAnkiServerUrl(): string {
  const config = getResolvedConfig().ankiConnect;
  if (config.proxy?.enabled) {
    const host = config.proxy.host || '127.0.0.1';
    const port = config.proxy.port || 8766;
    return `http://${host}:${port}`;
  }
  return config.url;
}

async function syncYomitanDefaultProfileAnkiServer(): Promise<void> {
  const targetUrl = getPreferredYomitanAnkiServerUrl().trim();
  if (!targetUrl || targetUrl === lastSyncedYomitanAnkiServer) {
    return;
  }

  const synced = await syncYomitanDefaultAnkiServerCore(
    targetUrl,
    {
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
    },
    {
      error: (message, ...args) => {
        logger.error(message, ...args);
      },
      info: (message, ...args) => {
        logger.info(message, ...args);
      },
    },
  );

  if (synced) {
    lastSyncedYomitanAnkiServer = targetUrl;
  }
}

function createOverlayWindow(kind: 'visible' | 'modal'): BrowserWindow {
  return createOverlayWindowHandler(kind);
}

function createModalWindow(): BrowserWindow {
  const existingWindow = overlayManager.getModalWindow();
  if (existingWindow && !existingWindow.isDestroyed()) {
    return existingWindow;
  }
  const window = createModalWindowHandler();
  overlayManager.setModalWindowBounds(getCurrentOverlayGeometry());
  return window;
}

function createMainWindow(): BrowserWindow {
  return createMainWindowHandler();
}
function resolveTrayIconPath(): string | null {
  return resolveTrayIconPathHandler();
}

function buildTrayMenu(): Menu {
  return buildTrayMenuHandler();
}

function ensureTray(): void {
  ensureTrayHandler();
}

function destroyTray(): void {
  destroyTrayHandler();
}

function initializeOverlayRuntime(): void {
  initializeOverlayRuntimeHandler();
  syncOverlayMpvSubtitleSuppression();
}

function openYomitanSettings(): void {
  openYomitanSettingsHandler();
}

const {
  getConfiguredShortcuts,
  registerGlobalShortcuts,
  refreshGlobalAndOverlayShortcuts,
  cancelPendingMultiCopy,
  startPendingMultiCopy,
  cancelPendingMineSentenceMultiple,
  startPendingMineSentenceMultiple,
  registerOverlayShortcuts,
  unregisterOverlayShortcuts,
  syncOverlayShortcuts,
  refreshOverlayShortcuts,
} = composeShortcutRuntimes({
  globalShortcuts: {
    getConfiguredShortcutsMainDeps: {
      getResolvedConfig: () => getResolvedConfig(),
      defaultConfig: DEFAULT_CONFIG,
      resolveConfiguredShortcuts,
    },
    buildRegisterGlobalShortcutsMainDeps: (getConfiguredShortcutsHandler) => ({
      getConfiguredShortcuts: () => getConfiguredShortcutsHandler(),
      registerGlobalShortcutsCore,
      toggleVisibleOverlay: () => toggleVisibleOverlay(),
      openYomitanSettings: () => openYomitanSettings(),
      isDev,
      getMainWindow: () => overlayManager.getMainWindow(),
    }),
    buildRefreshGlobalAndOverlayShortcutsMainDeps: (registerGlobalShortcutsHandler) => ({
      unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
      registerGlobalShortcuts: () => registerGlobalShortcutsHandler(),
      syncOverlayShortcuts: () => syncOverlayShortcuts(),
    }),
  },
  numericShortcutRuntimeMainDeps: {
    globalShortcut,
    showMpvOsd: (text) => showMpvOsd(text),
    setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
    clearTimer: (timer) => clearTimeout(timer),
  },
  numericSessions: {
    onMultiCopyDigit: (count) => handleMultiCopyDigit(count),
    onMineSentenceDigit: (count) => handleMineSentenceDigit(count),
  },
  overlayShortcutsRuntimeMainDeps: {
    overlayShortcutsRuntime,
  },
});

const { appendToMpvLog, flushMpvLog, showMpvOsd } = createMpvOsdRuntimeHandlers({
  appendToMpvLogMainDeps: {
    logPath: DEFAULT_MPV_LOG_PATH,
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

const cycleSecondarySubMode = createCycleSecondarySubModeRuntimeHandler({
  cycleSecondarySubModeMainDeps: {
    getSecondarySubMode: () => appState.secondarySubMode,
    setSecondarySubMode: (mode: SecondarySubMode) => {
      setSecondarySubMode(mode);
    },
    getLastSecondarySubToggleAtMs: () => appState.lastSecondarySubToggleAtMs,
    setLastSecondarySubToggleAtMs: (timestampMs: number) => {
      appState.lastSecondarySubToggleAtMs = timestampMs;
    },
    broadcastToOverlayWindows: (channel, mode) => {
      broadcastToOverlayWindows(channel, mode);
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
  },
  cycleSecondarySubMode: (deps) => cycleSecondarySubModeCore(deps),
});

function setSecondarySubMode(mode: SecondarySubMode): void {
  appState.secondarySubMode = mode;
}

function handleCycleSecondarySubMode(): void {
  cycleSecondarySubMode();
}

async function triggerSubsyncFromConfig(): Promise<void> {
  await subsyncRuntime.triggerFromConfig();
}

function handleMultiCopyDigit(count: number): void {
  handleMultiCopyDigitHandler(count);
}

function copyCurrentSubtitle(): void {
  copyCurrentSubtitleHandler();
}

const buildUpdateLastCardFromClipboardMainDepsHandler =
  createBuildUpdateLastCardFromClipboardMainDepsHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    readClipboardText: () => clipboard.readText(),
    showMpvOsd: (text) => showMpvOsd(text),
    updateLastCardFromClipboardCore,
  });
const updateLastCardFromClipboardMainDeps = buildUpdateLastCardFromClipboardMainDepsHandler();
const updateLastCardFromClipboardHandler = createUpdateLastCardFromClipboardHandler(
  updateLastCardFromClipboardMainDeps,
);

const buildRefreshKnownWordCacheMainDepsHandler = createBuildRefreshKnownWordCacheMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  missingIntegrationMessage: 'AnkiConnect integration not enabled',
});
const refreshKnownWordCacheMainDeps = buildRefreshKnownWordCacheMainDepsHandler();
const refreshKnownWordCacheHandler = createRefreshKnownWordCacheHandler(
  refreshKnownWordCacheMainDeps,
);

const buildTriggerFieldGroupingMainDepsHandler = createBuildTriggerFieldGroupingMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  showMpvOsd: (text) => showMpvOsd(text),
  triggerFieldGroupingCore,
});
const triggerFieldGroupingMainDeps = buildTriggerFieldGroupingMainDepsHandler();
const triggerFieldGroupingHandler = createTriggerFieldGroupingHandler(triggerFieldGroupingMainDeps);

const buildMarkLastCardAsAudioCardMainDepsHandler =
  createBuildMarkLastCardAsAudioCardMainDepsHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
    markLastCardAsAudioCardCore,
  });
const markLastCardAsAudioCardMainDeps = buildMarkLastCardAsAudioCardMainDepsHandler();
const markLastCardAsAudioCardHandler = createMarkLastCardAsAudioCardHandler(
  markLastCardAsAudioCardMainDeps,
);

const buildMineSentenceCardMainDepsHandler = createBuildMineSentenceCardMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  getMpvClient: () => appState.mpvClient,
  showMpvOsd: (text) => showMpvOsd(text),
  mineSentenceCardCore,
  recordCardsMined: (count) => {
    ensureImmersionTrackerStarted();
    appState.immersionTracker?.recordCardsMined(count);
  },
});
const mineSentenceCardHandler = createMineSentenceCardHandler(
  buildMineSentenceCardMainDepsHandler(),
);

const buildHandleMultiCopyDigitMainDepsHandler = createBuildHandleMultiCopyDigitMainDepsHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => showMpvOsd(text),
  handleMultiCopyDigitCore,
});
const handleMultiCopyDigitMainDeps = buildHandleMultiCopyDigitMainDepsHandler();
const handleMultiCopyDigitHandler = createHandleMultiCopyDigitHandler(handleMultiCopyDigitMainDeps);

const buildCopyCurrentSubtitleMainDepsHandler = createBuildCopyCurrentSubtitleMainDepsHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => showMpvOsd(text),
  copyCurrentSubtitleCore,
});
const copyCurrentSubtitleMainDeps = buildCopyCurrentSubtitleMainDepsHandler();
const copyCurrentSubtitleHandler = createCopyCurrentSubtitleHandler(copyCurrentSubtitleMainDeps);

const buildHandleMineSentenceDigitMainDepsHandler =
  createBuildHandleMineSentenceDigitMainDepsHandler({
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getAnkiIntegration: () => appState.ankiIntegration,
    getCurrentSecondarySubText: () => appState.mpvClient?.currentSecondarySubText || undefined,
    showMpvOsd: (text) => showMpvOsd(text),
    logError: (message, err) => {
      logger.error(message, err);
    },
    onCardsMined: (cards) => {
      ensureImmersionTrackerStarted();
      appState.immersionTracker?.recordCardsMined(cards);
    },
    handleMineSentenceDigitCore,
  });
const handleMineSentenceDigitMainDeps = buildHandleMineSentenceDigitMainDepsHandler();
const handleMineSentenceDigitHandler = createHandleMineSentenceDigitHandler(
  handleMineSentenceDigitMainDeps,
);
const {
  setVisibleOverlayVisible: setVisibleOverlayVisibleHandler,
  toggleVisibleOverlay: toggleVisibleOverlayHandler,
  setOverlayVisible: setOverlayVisibleHandler,
  toggleOverlay: toggleOverlayHandler,
} = createOverlayVisibilityRuntime({
  setVisibleOverlayVisibleDeps: {
    setVisibleOverlayVisibleCore,
    setVisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setVisibleOverlayVisible(nextVisible);
    },
    updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
  },
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
});

const buildHandleOverlayModalClosedMainDepsHandler =
  createBuildHandleOverlayModalClosedMainDepsHandler({
    handleOverlayModalClosedRuntime: (modal) => overlayModalRuntime.handleOverlayModalClosed(modal),
  });
const handleOverlayModalClosedMainDeps = buildHandleOverlayModalClosedMainDepsHandler();
const handleOverlayModalClosedHandler = createHandleOverlayModalClosedHandler(
  handleOverlayModalClosedMainDeps,
);

const buildAppendClipboardVideoToQueueMainDepsHandler =
  createBuildAppendClipboardVideoToQueueMainDepsHandler({
    appendClipboardVideoToQueueRuntime,
    getMpvClient: () => appState.mpvClient,
    readClipboardText: () => clipboard.readText(),
    showMpvOsd: (text) => showMpvOsd(text),
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
  });
const appendClipboardVideoToQueueMainDeps = buildAppendClipboardVideoToQueueMainDepsHandler();
const appendClipboardVideoToQueueHandler = createAppendClipboardVideoToQueueHandler(
  appendClipboardVideoToQueueMainDeps,
);

const shiftSubtitleDelayToAdjacentCueHandler = createShiftSubtitleDelayToAdjacentCueHandler({
  getMpvClient: () => appState.mpvClient,
  loadSubtitleSourceText: async (source) => {
    if (/^https?:\/\//i.test(source)) {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 4000);
      try {
        const response = await fetch(source, { signal: controller.signal });
        if (!response.ok) {
          throw new Error(`Failed to download subtitle source (${response.status})`);
        }
        return await response.text();
      } finally {
        clearTimeout(timeoutId);
      }
    }

    const filePath = source.startsWith('file://') ? decodeURI(new URL(source).pathname) : source;
    return fs.promises.readFile(filePath, 'utf8');
  },
  sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
  showMpvOsd: (text) => showMpvOsd(text),
});

const {
  handleMpvCommandFromIpc: handleMpvCommandFromIpcHandler,
  runSubsyncManualFromIpc: runSubsyncManualFromIpcHandler,
  registerIpcRuntimeHandlers,
} = composeIpcRuntimeHandlers({
  mpvCommandMainDeps: {
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
    shiftSubDelayToAdjacentSubtitle: (direction) =>
      shiftSubtitleDelayToAdjacentCueHandler(direction),
    sendMpvCommand: (rawCommand: (string | number)[]) =>
      sendMpvCommandRuntime(appState.mpvClient, rawCommand),
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
  },
  handleMpvCommandFromIpcRuntime,
  runSubsyncManualFromIpc: (request) => subsyncRuntime.runManualFromIpc(request),
  registration: {
    runtimeOptions: {
      getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
      showMpvOsd: (text: string) => showMpvOsd(text),
    },
    mainDeps: {
      getMainWindow: () => overlayManager.getMainWindow(),
      getVisibleOverlayVisibility: () => overlayManager.getVisibleOverlayVisible(),
      focusMainWindow: () => {
        const mainWindow = overlayManager.getMainWindow();
        if (!mainWindow || mainWindow.isDestroyed()) return;
        if (!mainWindow.isFocused()) {
          mainWindow.focus();
        }
      },
      onOverlayModalClosed: (modal) => {
        handleOverlayModalClosed(modal);
      },
      onOverlayModalOpened: (modal) => {
        overlayModalRuntime.notifyOverlayModalOpened(modal);
      },
      openYomitanSettings: () => openYomitanSettings(),
      quitApp: () => app.quit(),
      toggleVisibleOverlay: () => toggleVisibleOverlay(),
      tokenizeCurrentSubtitle: () => tokenizeSubtitle(appState.currentSubText),
      getCurrentSubtitleRaw: () => appState.currentSubText,
      getCurrentSubtitleAss: () => appState.currentSubAssText,
      getPlaybackPaused: () => appState.playbackPaused,
      getSubtitlePosition: () => loadSubtitlePosition(),
      getSubtitleStyle: () => {
        const resolvedConfig = getResolvedConfig();
        return resolveSubtitleStyleForRenderer(resolvedConfig);
      },
      saveSubtitlePosition: (position) => saveSubtitlePosition(position),
      getMecabTokenizer: () => appState.mecabTokenizer,
      getKeybindings: () => appState.keybindings,
      getConfiguredShortcuts: () => getConfiguredShortcuts(),
      getSecondarySubMode: () => appState.secondarySubMode,
      getMpvClient: () => appState.mpvClient,
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
    registerIpcRuntimeServices,
  },
});
const createCliCommandContextHandler = createCliCommandContextFactory({
  appState,
  setLogLevel: (level) => setLogLevel(level, 'cli'),
  texthookerService,
  getResolvedConfig: () => getResolvedConfig(),
  openExternal: (url: string) => shell.openExternal(url),
  logBrowserOpenError: (url: string, error: unknown) =>
    logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
  showMpvOsd: (text: string) => showMpvOsd(text),
  initializeOverlayRuntime: () => initializeOverlayRuntime(),
  toggleVisibleOverlay: () => toggleVisibleOverlay(),
  setVisibleOverlayVisible: (visible: boolean) => setVisibleOverlayVisible(visible),
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
  openAnilistSetupWindow: () => openAnilistSetupWindow(),
  openJellyfinSetupWindow: () => openJellyfinSetupWindow(),
  getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
  processNextAnilistRetryUpdate: () => processNextAnilistRetryUpdate(),
  runJellyfinCommand: (argsFromCommand: CliArgs) => runJellyfinCommand(argsFromCommand),
  openYomitanSettings: () => openYomitanSettings(),
  cycleSecondarySubMode: () => handleCycleSecondarySubMode(),
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
const {
  createOverlayWindow: createOverlayWindowHandler,
  createMainWindow: createMainWindowHandler,
  createModalWindow: createModalWindowHandler,
} = createOverlayWindowRuntimeHandlers<BrowserWindow>({
  createOverlayWindowDeps: {
    createOverlayWindowCore: (kind, options) => createOverlayWindowCore(kind, options),
    isDev,
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
    onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    setOverlayDebugVisualizationEnabled: (enabled) => setOverlayDebugVisualizationEnabled(enabled),
    isOverlayVisible: (windowKind) =>
      windowKind === 'visible' ? overlayManager.getVisibleOverlayVisible() : false,
    tryHandleOverlayShortcutLocalFallback: (input) =>
      overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(input),
    onWindowClosed: (windowKind) => {
      if (windowKind === 'visible') {
        overlayManager.setMainWindow(null);
      } else {
        overlayManager.setModalWindow(null);
      }
    },
  },
  setMainWindow: (window) => overlayManager.setMainWindow(window),
  setModalWindow: (window) => overlayManager.setModalWindow(window),
});
const {
  resolveTrayIconPath: resolveTrayIconPathHandler,
  buildTrayMenu: buildTrayMenuHandler,
  ensureTray: ensureTrayHandler,
  destroyTray: destroyTrayHandler,
} = createTrayRuntimeHandlers({
  resolveTrayIconPathDeps: {
    resolveTrayIconPathRuntime,
    platform: process.platform,
    resourcesPath: process.resourcesPath,
    appPath: app.getAppPath(),
    dirname: __dirname,
    joinPath: (...parts) => path.join(...parts),
    fileExists: (candidate) => fs.existsSync(candidate),
  },
  buildTrayMenuTemplateDeps: {
    buildTrayMenuTemplateRuntime,
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    openYomitanSettings: () => openYomitanSettings(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    openJellyfinSetupWindow: () => openJellyfinSetupWindow(),
    openAnilistSetupWindow: () => openAnilistSetupWindow(),
    quitApp: () => app.quit(),
  },
  ensureTrayDeps: {
    getTray: () => appTray,
    setTray: (tray) => {
      appTray = tray as Tray | null;
    },
    createImageFromPath: (iconPath) => nativeImage.createFromPath(iconPath),
    createEmptyImage: () => nativeImage.createEmpty(),
    createTray: (icon) => new Tray(icon as ConstructorParameters<typeof Tray>[0]),
    trayTooltip: TRAY_TOOLTIP,
    platform: process.platform,
    logWarn: (message) => logger.warn(message),
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  },
  destroyTrayDeps: {
    getTray: () => appTray,
    setTray: (tray) => {
      appTray = tray as Tray | null;
    },
  },
  buildMenuFromTemplate: (template) => Menu.buildFromTemplate(template),
});
const yomitanExtensionRuntime = createYomitanExtensionRuntime({
  loadYomitanExtensionCore,
  userDataPath: USER_DATA_PATH,
  getYomitanParserWindow: () => appState.yomitanParserWindow,
  setYomitanParserWindow: (window) => {
    appState.yomitanParserWindow = window as BrowserWindow | null;
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
  getYomitanExtension: () => appState.yomitanExt,
  getLoadInFlight: () => yomitanLoadInFlight,
  setLoadInFlight: (promise) => {
    yomitanLoadInFlight = promise;
  },
});
const { initializeOverlayRuntime: initializeOverlayRuntimeHandler } =
  runtimeRegistry.overlay.createOverlayRuntimeBootstrapHandlers({
    initializeOverlayRuntimeMainDeps: {
      appState,
      overlayManager: {
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () =>
          overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
      },
      overlayShortcutsRuntime: {
        syncOverlayShortcuts: () => overlayShortcutsRuntime.syncOverlayShortcuts(),
      },
      createMainWindow: () => createMainWindow(),
      registerGlobalShortcuts: () => registerGlobalShortcuts(),
      updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
        updateVisibleOverlayBounds(geometry),
      getOverlayWindows: () => getOverlayWindows(),
      getResolvedConfig: () => getResolvedConfig(),
      showDesktopNotification,
      createFieldGroupingCallback: () => createFieldGroupingCallback(),
      getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
    },
    initializeOverlayRuntimeBootstrapDeps: {
      isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
      initializeOverlayRuntimeCore,
      setOverlayRuntimeInitialized: (initialized) => {
        appState.overlayRuntimeInitialized = initialized;
      },
      startBackgroundWarmups: () => startBackgroundWarmups(),
    },
  });
const { openYomitanSettings: openYomitanSettingsHandler } = createYomitanSettingsRuntime({
  ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded(),
  openYomitanSettingsWindow: ({ yomitanExt, getExistingWindow, setWindow }) => {
    openYomitanSettingsWindow({
      yomitanExt: yomitanExt as Extension,
      getExistingWindow: () => getExistingWindow() as BrowserWindow | null,
      setWindow: (window) => setWindow(window as BrowserWindow | null),
      onWindowClosed: () => {
        if (appState.yomitanParserWindow) {
          clearYomitanParserCachesForWindow(appState.yomitanParserWindow);
        }
      },
    });
  },
  getExistingWindow: () => appState.yomitanSettingsWindow,
  setWindow: (window) => {
    appState.yomitanSettingsWindow = window as BrowserWindow | null;
  },
  logWarn: (message) => logger.warn(message),
  logError: (message, error) => logger.error(message, error),
});

async function updateLastCardFromClipboard(): Promise<void> {
  await updateLastCardFromClipboardHandler();
}

async function refreshKnownWordCache(): Promise<void> {
  await refreshKnownWordCacheHandler();
}

async function triggerFieldGrouping(): Promise<void> {
  await triggerFieldGroupingHandler();
}

async function markLastCardAsAudioCard(): Promise<void> {
  await markLastCardAsAudioCardHandler();
}

async function mineSentenceCard(): Promise<void> {
  await mineSentenceCardHandler();
}

function handleMineSentenceDigit(count: number): void {
  handleMineSentenceDigitHandler(count);
}

function ensureOverlayWindowsReadyForVisibilityActions(): void {
  if (!appState.overlayRuntimeInitialized) {
    initializeOverlayRuntime();
    return;
  }

  const mainWindow = overlayManager.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed()) {
    createMainWindow();
  }
}

function setVisibleOverlayVisible(visible: boolean): void {
  ensureOverlayWindowsReadyForVisibilityActions();
  if (visible) {
    void ensureOverlayMpvSubtitlesHidden();
  }
  setVisibleOverlayVisibleHandler(visible);
  syncOverlayMpvSubtitleSuppression();
}

function toggleVisibleOverlay(): void {
  ensureOverlayWindowsReadyForVisibilityActions();
  if (!overlayManager.getVisibleOverlayVisible()) {
    void ensureOverlayMpvSubtitlesHidden();
  }
  toggleVisibleOverlayHandler();
  syncOverlayMpvSubtitleSuppression();
}
function setOverlayVisible(visible: boolean): void {
  if (visible) {
    void ensureOverlayMpvSubtitlesHidden();
  }
  setOverlayVisibleHandler(visible);
  syncOverlayMpvSubtitleSuppression();
}
function toggleOverlay(): void {
  if (!overlayManager.getVisibleOverlayVisible()) {
    void ensureOverlayMpvSubtitlesHidden();
  }
  toggleOverlayHandler();
  syncOverlayMpvSubtitleSuppression();
}
function handleOverlayModalClosed(modal: OverlayHostedModal): void {
  handleOverlayModalClosedHandler(modal);
}

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcHandler(command);
}

async function runSubsyncManualFromIpc(request: SubsyncManualRunRequest): Promise<SubsyncResult> {
  return runSubsyncManualFromIpcHandler(request) as Promise<SubsyncResult>;
}

function appendClipboardVideoToQueue(): { ok: boolean; message: string } {
  return appendClipboardVideoToQueueHandler();
}

registerIpcRuntimeHandlers();
