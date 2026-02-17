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
} from "electron";

protocol.registerSchemesAsPrivileged([
  {
    scheme: "chrome-extension",
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
      bypassCSP: true,
    },
  },
]);

import * as path from "path";
import * as os from "os";
import * as fs from "fs";
import { MecabTokenizer } from "./mecab-tokenizer";
import type {
  JimakuApiResponse,
  JimakuLanguagePreference,
  SubtitleData,
  SubtitlePosition,
  WindowGeometry,
  SecondarySubMode,
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncResult,
  KikuFieldGroupingChoice,
  RuntimeOptionState,
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
} from "./types";
import { SubtitleTimingTracker } from "./subtitle-timing-tracker";
import { AnkiIntegration } from "./anki-integration";
import { RuntimeOptionsManager } from "./runtime-options";
import {
  downloadToFile,
  isRemoteMediaPath,
  parseMediaInfo,
} from "./jimaku/utils";
import {
  getSubsyncConfig,
} from "./subsync/utils";
import { createLogger, setLogLevel, type LogLevelSource } from "./logger";
import {
  parseArgs,
  shouldStartApp,
} from "./cli/args";
import type { CliArgs, CliCommandSource } from "./cli/args";
import { printHelp } from "./cli/help";
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
  generateDefaultConfigFile,
  resolveConfiguredShortcuts,
  resolveKeybindings,
  showDesktopNotification,
} from "./core/utils";
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
  createOverlayWindow as createOverlayWindowCore,
  createTokenizerDepsRuntime,
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  enforceOverlayLayerOrder as enforceOverlayLayerOrderCore,
  ensureOverlayWindowLevel as ensureOverlayWindowLevelCore,
  getInitialInvisibleOverlayVisibility as getInitialInvisibleOverlayVisibilityCore,
  getJimakuLanguagePreference as getJimakuLanguagePreferenceCore,
  getJimakuMaxEntryResults as getJimakuMaxEntryResultsCore,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  hasMpvWebsocketPlugin,
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  isAutoUpdateEnabledRuntime as isAutoUpdateEnabledRuntimeCore,
  jimakuFetchJson as jimakuFetchJsonCore,
  loadSubtitlePosition as loadSubtitlePositionCore,
  loadYomitanExtension as loadYomitanExtensionCore,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
  mineSentenceCard as mineSentenceCardCore,
  ImmersionTrackerService,
  openYomitanSettingsWindow,
  playNextSubtitleRuntime,
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  replayCurrentSubtitleRuntime,
  resolveJimakuApiKey as resolveJimakuApiKeyCore,
  runStartupBootstrapRuntime,
  saveSubtitlePosition as saveSubtitlePositionCore,
  sendMpvCommandRuntime,
  setInvisibleOverlayVisible as setInvisibleOverlayVisibleCore,
  setMpvSubVisibilityRuntime,
  setOverlayDebugVisualizationEnabledRuntime,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  shouldAutoInitializeOverlayRuntimeFromConfig as shouldAutoInitializeOverlayRuntimeFromConfigCore,
  shouldBindVisibleOverlayToMpvSubVisibility as shouldBindVisibleOverlayToMpvSubVisibilityCore,
  showMpvOsdRuntime,
  tokenizeSubtitle as tokenizeSubtitleCore,
  triggerFieldGrouping as triggerFieldGroupingCore,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
} from "./core/services";
import {
  guessAnilistMediaInfo,
  type AnilistMediaGuess,
  updateAnilistPostWatchProgress,
} from "./core/services/anilist/anilist-updater";
import { createAnilistTokenStore } from "./core/services/anilist/anilist-token-store";
import { createAnilistUpdateQueue } from "./core/services/anilist/anilist-update-queue";
import { applyRuntimeOptionResultRuntime } from "./core/services/runtime-options-ipc";
import {
  createAppReadyRuntimeRunner,
} from "./main/app-lifecycle";
import { handleMpvCommandFromIpcRuntime } from "./main/ipc-mpv-command";
import {
  registerIpcRuntimeServices,
} from "./main/ipc-runtime";
import {
  createAnkiJimakuIpcRuntimeServiceDeps,
} from "./main/dependencies";
import {
  handleCliCommandRuntimeServiceWithContext,
} from "./main/cli-runtime";
import {
  runSubsyncManualFromIpcRuntime,
  triggerSubsyncFromConfigRuntime,
  createSubsyncRuntimeServiceInputFromState,
} from "./main/subsync-runtime";
import {
  createOverlayModalRuntimeService,
  type OverlayHostedModal,
} from "./main/overlay-runtime";
import {
  createOverlayShortcutsRuntimeService,
} from "./main/overlay-shortcuts-runtime";
import {
  createJlptDictionaryRuntimeService,
  getJlptDictionarySearchPaths,
} from "./main/jlpt-runtime";
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from "./main/frequency-dictionary-runtime";
import { createMediaRuntimeService } from "./main/media-runtime";
import { createOverlayVisibilityRuntimeService } from "./main/overlay-visibility-runtime";
import {
  type AppState,
  applyStartupState,
  createAppState,
} from "./main/state";
import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from "./main/anilist-url-guard";
import { createStartupBootstrapRuntimeDeps } from "./main/startup";
import { createAppLifecycleRuntimeRunner } from "./main/startup-lifecycle";
import {
  ConfigService,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
} from "./config";

if (process.platform === "linux") {
  app.commandLine.appendSwitch("enable-features", "GlobalShortcutsPortal");
}

const DEFAULT_TEXTHOOKER_PORT = 5174;
const DEFAULT_MPV_LOG_FILE = path.join(
  os.homedir(),
  ".cache",
  "SubMiner",
  "mp.log",
);
const ANILIST_SETUP_CLIENT_ID_URL = "https://anilist.co/api/v2/oauth/authorize";
const ANILIST_SETUP_RESPONSE_TYPE = "token";
const ANILIST_DEFAULT_CLIENT_ID = "36084";
const ANILIST_REDIRECT_URI = "https://anilist.subminer.moe/";
const ANILIST_DEVELOPER_SETTINGS_URL = "https://anilist.co/settings/developer";
const ANILIST_UPDATE_MIN_WATCH_RATIO = 0.85;
const ANILIST_UPDATE_MIN_WATCH_SECONDS = 10 * 60;
const ANILIST_DURATION_RETRY_INTERVAL_MS = 15_000;
const ANILIST_MAX_ATTEMPTED_UPDATE_KEYS = 1000;
const ANILIST_TOKEN_STORE_FILE = "anilist-token-store.json";
const ANILIST_RETRY_QUEUE_FILE = "anilist-retry-queue.json";

let anilistCurrentMediaKey: string | null = null;
let anilistCurrentMediaDurationSec: number | null = null;
let anilistCurrentMediaGuess: AnilistMediaGuess | null = null;
let anilistCurrentMediaGuessPromise: Promise<AnilistMediaGuess | null> | null = null;
let anilistLastDurationProbeAtMs = 0;
let anilistUpdateInFlight = false;
const anilistAttemptedUpdateKeys = new Set<string>();
let anilistCachedAccessToken: string | null = null;

function resolveConfigDir(): string {
  const xdgConfigHome = process.env.XDG_CONFIG_HOME?.trim();
  const baseDirs = Array.from(
    new Set([
      xdgConfigHome || path.join(os.homedir(), ".config"),
      path.join(os.homedir(), ".config"),
    ]),
  );
  const appNames = ["SubMiner", "subminer"];

  for (const baseDir of baseDirs) {
    for (const appName of appNames) {
      const dir = path.join(baseDir, appName);
      if (
        fs.existsSync(path.join(dir, "config.jsonc")) ||
        fs.existsSync(path.join(dir, "config.json"))
      ) {
        return dir;
      }
    }
  }

  for (const baseDir of baseDirs) {
    for (const appName of appNames) {
      const dir = path.join(baseDir, appName);
      if (fs.existsSync(dir)) {
        return dir;
      }
    }
  }

  return path.join(baseDirs[0], "SubMiner");
}

const CONFIG_DIR = resolveConfigDir();
const USER_DATA_PATH = CONFIG_DIR;
const DEFAULT_MPV_LOG_PATH = process.env.SUBMINER_MPV_LOG?.trim() || DEFAULT_MPV_LOG_FILE;
const DEFAULT_IMMERSION_DB_PATH = path.join(USER_DATA_PATH, "immersion.sqlite");
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
const isDev =
  process.argv.includes("--dev") || process.argv.includes("--debug");
const texthookerService = new Texthooker();
const subtitleWsService = new SubtitleWebSocket();
const logger = createLogger("main");
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
    logger.error("No running instance. Use --start to launch the app.");
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
  if (process.platform === "win32") {
    return "\\\\.\\pipe\\subminer-socket";
  }
  return "/tmp/subminer-socket";
}

if (!fs.existsSync(USER_DATA_PATH)) {
  fs.mkdirSync(USER_DATA_PATH, { recursive: true });
}
app.setPath("userData", USER_DATA_PATH);

process.on("SIGINT", () => {
  app.quit();
});
process.on("SIGTERM", () => {
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
    sendToActiveOverlayWindow("jimaku:open", undefined, {
      restoreOnModalClose: "jimaku",
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

const jlptDictionaryRuntime = createJlptDictionaryRuntimeService({
  isJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
  getSearchPaths: () =>
    getJlptDictionarySearchPaths({
      getDictionaryRoots: () => [
        path.join(__dirname, "..", "..", "vendor", "yomitan-jlpt-vocab"),
        path.join(app.getAppPath(), "vendor", "yomitan-jlpt-vocab"),
        path.join(process.resourcesPath, "yomitan-jlpt-vocab"),
        path.join(process.resourcesPath, "app.asar", "vendor", "yomitan-jlpt-vocab"),
        USER_DATA_PATH,
        app.getPath("userData"),
        path.join(os.homedir(), ".config", "SubMiner"),
        path.join(os.homedir(), ".config", "subminer"),
        path.join(os.homedir(), "Library", "Application Support", "SubMiner"),
        path.join(os.homedir(), "Library", "Application Support", "subminer"),
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
  isFrequencyDictionaryEnabled: () =>
    getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
  getSearchPaths: () =>
    getFrequencyDictionarySearchPaths({
      getDictionaryRoots: () => [
        path.join(__dirname, "..", "..", "vendor", "jiten_freq_global"),
        path.join(__dirname, "..", "..", "vendor", "frequency-dictionary"),
        path.join(app.getAppPath(), "vendor", "jiten_freq_global"),
        path.join(app.getAppPath(), "vendor", "frequency-dictionary"),
        path.join(process.resourcesPath, "jiten_freq_global"),
        path.join(process.resourcesPath, "frequency-dictionary"),
        path.join(process.resourcesPath, "app.asar", "vendor", "jiten_freq_global"),
        path.join(process.resourcesPath, "app.asar", "vendor", "frequency-dictionary"),
        USER_DATA_PATH,
        app.getPath("userData"),
        path.join(os.homedir(), ".config", "SubMiner"),
        path.join(os.homedir(), ".config", "subminer"),
        path.join(os.homedir(), "Library", "Application Support", "SubMiner"),
        path.join(os.homedir(), "Library", "Application Support", "subminer"),
        process.cwd(),
      ].filter((dictionaryRoot) => dictionaryRoot),
      getSourcePath: () =>
        getResolvedConfig().subtitleStyle.frequencyDictionary.sourcePath,
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
    return overlayModalRuntime.sendToActiveOverlayWindow(
      channel,
      payload,
      runtimeOptions,
    );
  },
});
const createFieldGroupingCallback =
  fieldGroupingOverlayRuntime.createFieldGroupingCallback;

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, "subtitle-positions");

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
    broadcastToOverlayWindows("subtitle-position:set", position);
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
  updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
    updateVisibleOverlayBounds(geometry),
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

function getRuntimeOptionsState(): RuntimeOptionState[] { if (!appState.runtimeOptionsManager) return []; return appState.runtimeOptionsManager.listOptions(); }

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
  return overlayModalRuntime.sendToActiveOverlayWindow(
    channel,
    payload,
    runtimeOptions,
  );
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

function getResolvedConfig() { return configService.getConfig(); }
function getConfiguredImmersionDbPath(): string {
  const configuredDbPath = getResolvedConfig().immersionTracking?.dbPath?.trim();
  return configuredDbPath
    ? configuredDbPath
    : DEFAULT_IMMERSION_DB_PATH;
}

let isImmersionTrackerMediaSeedInProgress = false;

type ImmersionMediaState = {
  path: string | null;
  title: string | null;
};

async function readMpvPropertyAsString(
  mpvClient: MpvIpcClient | null | undefined,
  propertyName: string,
): Promise<string | null> {
  if (!mpvClient) {
    return null;
  }
  try {
    const value = await mpvClient.requestProperty(propertyName);
    return typeof value === "string" ? value.trim() || null : null;
  } catch {
    return null;
  }
}

async function getCurrentMpvMediaStateForTracker(): Promise<ImmersionMediaState> {
  const statePath = appState.currentMediaPath?.trim() || null;
  if (statePath) {
    return {
      path: statePath,
      title: appState.currentMediaTitle?.trim() || null,
    };
  }

  const mpvClient = appState.mpvClient;
  const trackedPath = mpvClient?.currentVideoPath?.trim() || null;
  if (trackedPath) {
    return {
      path: trackedPath,
      title: appState.currentMediaTitle?.trim() || null,
    };
  }

  const [pathFromProperty, filenameFromProperty, titleFromProperty] =
    await Promise.all([
      readMpvPropertyAsString(mpvClient, "path"),
      readMpvPropertyAsString(mpvClient, "filename"),
      readMpvPropertyAsString(mpvClient, "media-title"),
    ]);

  const resolvedPath = pathFromProperty || filenameFromProperty || null;
  const resolvedTitle = appState.currentMediaTitle?.trim() || titleFromProperty || null;

  return {
    path: resolvedPath,
    title: resolvedTitle,
  };
}

function getInitialInvisibleOverlayVisibility(): boolean {
  return getInitialInvisibleOverlayVisibilityCore(
    getResolvedConfig(),
    process.platform,
  );
}

function shouldAutoInitializeOverlayRuntimeFromConfig(): boolean {
  return shouldAutoInitializeOverlayRuntimeFromConfigCore(getResolvedConfig());
}

function shouldBindVisibleOverlayToMpvSubVisibility(): boolean {
  return shouldBindVisibleOverlayToMpvSubVisibilityCore(getResolvedConfig());
}

function isAutoUpdateEnabledRuntime(): boolean {
  return isAutoUpdateEnabledRuntimeCore(
    getResolvedConfig(),
    appState.runtimeOptionsManager,
  );
}

function getJimakuLanguagePreference(): JimakuLanguagePreference { return getJimakuLanguagePreferenceCore(() => getResolvedConfig(), DEFAULT_CONFIG.jimaku.languagePreference); }

function getJimakuMaxEntryResults(): number { return getJimakuMaxEntryResultsCore(() => getResolvedConfig(), DEFAULT_CONFIG.jimaku.maxEntryResults); }

async function resolveJimakuApiKey(): Promise<string | null> { return resolveJimakuApiKeyCore(() => getResolvedConfig()); }

function seedImmersionTrackerFromCurrentMedia(): void {
  const tracker = appState.immersionTracker;
  if (!tracker) {
    logger.debug("Immersion tracker seeding skipped: tracker not initialized.");
    return;
  }
  if (isImmersionTrackerMediaSeedInProgress) {
    logger.debug(
      "Immersion tracker seeding already in progress; skipping duplicate call.",
    );
    return;
  }
  logger.debug("Starting immersion tracker media-state seed loop.");
  isImmersionTrackerMediaSeedInProgress = true;

  void (async () => {
    const waitMs = 250;
    const attempts = 120;
    for (let attempt = 0; attempt < attempts; attempt += 1) {
      const mediaState = await getCurrentMpvMediaStateForTracker();
      if (mediaState.path) {
        logger.info(
          `Seeded immersion tracker media state at attempt ${attempt + 1}/${attempts}: ` +
            `${mediaState.path}`,
        );
        tracker.handleMediaChange(mediaState.path, mediaState.title);
        return;
      }

      const mpvClient = appState.mpvClient;
      if (!mpvClient || !mpvClient.connected) {
        await new Promise((resolve) => setTimeout(resolve, waitMs));
        continue;
      }
      if (attempt < attempts - 1) {
        await new Promise((resolve) => setTimeout(resolve, waitMs));
      }
    }

    logger.info(
      "Immersion tracker seed failed: media path still unavailable after startup warmup",
    );
  })().finally(() => {
    isImmersionTrackerMediaSeedInProgress = false;
  });
}

function syncImmersionTrackerFromCurrentMediaState(): void {
  const tracker = appState.immersionTracker;
  if (!tracker) {
    logger.debug(
      "Immersion tracker sync skipped: tracker not initialized yet.",
    );
    return;
  }

  const pathFromState = appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim();
  if (pathFromState) {
    logger.debug(
      "Immersion tracker sync using path from current media state.",
    );
    tracker.handleMediaChange(pathFromState, appState.currentMediaTitle);
    return;
  }

  if (!isImmersionTrackerMediaSeedInProgress) {
    logger.debug(
      "Immersion tracker sync did not find media path; starting seed loop.",
    );
    seedImmersionTrackerFromCurrentMedia();
  } else {
    logger.debug(
      "Immersion tracker sync found seed loop already running.",
    );
  }
}

async function jimakuFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined> = {},
): Promise<JimakuApiResponse<T>> {
  return jimakuFetchJsonCore<T>(endpoint, query, {
    getResolvedConfig: () => getResolvedConfig(),
    defaultBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
    defaultMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
    defaultLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  });
}

function setAnilistClientSecretState(partial: Partial<AppState["anilistClientSecretState"]>): void {
  appState.anilistClientSecretState = {
    ...appState.anilistClientSecretState,
    ...partial,
  };
}

function refreshAnilistRetryQueueState(): void {
  appState.anilistRetryQueueState = {
    ...appState.anilistRetryQueueState,
    ...anilistUpdateQueue.getSnapshot(),
  };
}

function getAnilistStatusSnapshot() {
  return {
    tokenStatus: appState.anilistClientSecretState.status,
    tokenSource: appState.anilistClientSecretState.source,
    tokenMessage: appState.anilistClientSecretState.message,
    tokenResolvedAt: appState.anilistClientSecretState.resolvedAt,
    tokenErrorAt: appState.anilistClientSecretState.errorAt,
    queuePending: appState.anilistRetryQueueState.pending,
    queueReady: appState.anilistRetryQueueState.ready,
    queueDeadLetter: appState.anilistRetryQueueState.deadLetter,
    queueLastAttemptAt: appState.anilistRetryQueueState.lastAttemptAt,
    queueLastError: appState.anilistRetryQueueState.lastError,
  };
}

function getAnilistQueueStatusSnapshot() {
  refreshAnilistRetryQueueState();
  return {
    pending: appState.anilistRetryQueueState.pending,
    ready: appState.anilistRetryQueueState.ready,
    deadLetter: appState.anilistRetryQueueState.deadLetter,
    lastAttemptAt: appState.anilistRetryQueueState.lastAttemptAt,
    lastError: appState.anilistRetryQueueState.lastError,
  };
}

function clearAnilistTokenState(): void {
  anilistTokenStore.clearToken();
  anilistCachedAccessToken = null;
  setAnilistClientSecretState({
    status: "not_checked",
    source: "none",
    message: "stored token cleared",
    resolvedAt: null,
    errorAt: null,
  });
}

function isAnilistTrackingEnabled(resolved: ResolvedConfig): boolean {
  return resolved.anilist.enabled;
}

function buildAnilistSetupUrl(): string {
  const authorizeUrl = new URL(ANILIST_SETUP_CLIENT_ID_URL);
  authorizeUrl.searchParams.set("client_id", ANILIST_DEFAULT_CLIENT_ID);
  authorizeUrl.searchParams.set("response_type", ANILIST_SETUP_RESPONSE_TYPE);
  authorizeUrl.searchParams.set("redirect_uri", ANILIST_REDIRECT_URI);
  return authorizeUrl.toString();
}

function openAnilistSetupInBrowser(): void {
  const authorizeUrl = buildAnilistSetupUrl();
  void shell.openExternal(authorizeUrl).catch((error) => {
    logger.error("Failed to open AniList authorize URL in browser", error);
  });
}

function loadAnilistSetupFallback(setupWindow: BrowserWindow, reason: string): void {
  const authorizeUrl = buildAnilistSetupUrl();
  const fallbackHtml = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>AniList Setup</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 0; padding: 24px; background: #0b1020; color: #e5e7eb; }
    h1 { margin: 0 0 12px; font-size: 22px; }
    p { margin: 10px 0; line-height: 1.45; color: #cbd5e1; }
    a { color: #93c5fd; word-break: break-all; }
    .box { background: #111827; border: 1px solid #1f2937; border-radius: 10px; padding: 16px; }
    .reason { color: #fca5a5; }
  </style>
</head>
<body>
  <h1>AniList Setup</h1>
  <div class="box">
    <p class="reason">Embedded AniList page did not render: ${reason}</p>
    <p>We attempted to open the authorize URL in your default browser automatically.</p>
    <p>Use one of these links to continue setup:</p>
    <p><a href="${authorizeUrl}">${authorizeUrl}</a></p>
    <p><a href="${ANILIST_DEVELOPER_SETTINGS_URL}">${ANILIST_DEVELOPER_SETTINGS_URL}</a></p>
    <p>After login/authorization, copy the token into <code>anilist.accessToken</code>.</p>
  </div>
</body>
</html>`;
  void setupWindow.loadURL(
    `data:text/html;charset=utf-8,${encodeURIComponent(fallbackHtml)}`,
  );
}

function openAnilistSetupWindow(): void {
  if (appState.anilistSetupWindow) {
    appState.anilistSetupWindow.focus();
    return;
  }

  const setupWindow = new BrowserWindow({
    width: 1000,
    height: 760,
    title: "Anilist Setup",
    show: true,
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  setupWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (!isAllowedAnilistExternalUrl(url)) {
      logger.warn("Blocked unsafe AniList setup external URL", { url });
      return { action: "deny" };
    }
    void shell.openExternal(url);
    return { action: "deny" };
  });
  setupWindow.webContents.on("will-navigate", (event, url) => {
    if (isAllowedAnilistSetupNavigationUrl(url)) {
      return;
    }
    event.preventDefault();
    logger.warn("Blocked unsafe AniList setup navigation URL", { url });
  });

  setupWindow.webContents.on(
    "did-fail-load",
    (_event, errorCode, errorDescription, validatedURL) => {
      logger.error("AniList setup window failed to load", {
        errorCode,
        errorDescription,
        validatedURL,
      });
      openAnilistSetupInBrowser();
      if (!setupWindow.isDestroyed()) {
        loadAnilistSetupFallback(
          setupWindow,
          `${errorDescription} (${errorCode})`,
        );
      }
    },
  );

  setupWindow.webContents.on("did-finish-load", () => {
    const loadedUrl = setupWindow.webContents.getURL();
    if (!loadedUrl || loadedUrl === "about:blank") {
      logger.warn("AniList setup loaded a blank page; using fallback");
      openAnilistSetupInBrowser();
      if (!setupWindow.isDestroyed()) {
        loadAnilistSetupFallback(setupWindow, "blank page");
      }
    }
  });

  void setupWindow.loadURL(buildAnilistSetupUrl()).catch((error) => {
    logger.error("AniList setup loadURL rejected", error);
    openAnilistSetupInBrowser();
    if (!setupWindow.isDestroyed()) {
      loadAnilistSetupFallback(
        setupWindow,
        error instanceof Error ? error.message : String(error),
      );
    }
  });

  setupWindow.on("closed", () => {
    appState.anilistSetupWindow = null;
    appState.anilistSetupPageOpened = false;
  });

  appState.anilistSetupWindow = setupWindow;
  appState.anilistSetupPageOpened = true;
}

async function refreshAnilistClientSecretState(options?: { force?: boolean }): Promise<string | null> {
  const resolved = getResolvedConfig();
  const now = Date.now();
  if (!isAnilistTrackingEnabled(resolved)) {
    anilistCachedAccessToken = null;
    setAnilistClientSecretState({
      status: "not_checked",
      source: "none",
      message: "anilist tracking disabled",
      resolvedAt: null,
      errorAt: null,
    });
    appState.anilistSetupPageOpened = false;
    return null;
  }
  const rawAccessToken = resolved.anilist.accessToken.trim();
  if (rawAccessToken.length > 0) {
    if (options?.force || rawAccessToken !== anilistCachedAccessToken) {
      anilistTokenStore.saveToken(rawAccessToken);
    }
    anilistCachedAccessToken = rawAccessToken;
    setAnilistClientSecretState({
      status: "resolved",
      source: "literal",
      message: "using configured anilist.accessToken",
      resolvedAt: now,
      errorAt: null,
    });
    appState.anilistSetupPageOpened = false;
    return rawAccessToken;
  }

  if (!options?.force && anilistCachedAccessToken && anilistCachedAccessToken.length > 0) {
    return anilistCachedAccessToken;
  }

  const storedToken = anilistTokenStore.loadToken()?.trim() ?? "";
  if (storedToken.length > 0) {
    anilistCachedAccessToken = storedToken;
    setAnilistClientSecretState({
      status: "resolved",
      source: "stored",
      message: "using stored anilist access token",
      resolvedAt: now,
      errorAt: null,
    });
    appState.anilistSetupPageOpened = false;
    return storedToken;
  }

  anilistCachedAccessToken = null;
  setAnilistClientSecretState({
    status: "error",
    source: "none",
    message: "cannot authenticate without anilist.accessToken",
    resolvedAt: null,
    errorAt: now,
  });
  if (
    isAnilistTrackingEnabled(resolved) &&
    !appState.anilistSetupPageOpened
  ) {
    openAnilistSetupWindow();
  }
  return null;
}

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

async function maybeProbeAnilistDuration(mediaKey: string): Promise<number | null> {
  if (anilistCurrentMediaKey !== mediaKey) {
    return null;
  }
  if (
    typeof anilistCurrentMediaDurationSec === "number" &&
    anilistCurrentMediaDurationSec > 0
  ) {
    return anilistCurrentMediaDurationSec;
  }
  const now = Date.now();
  if (now - anilistLastDurationProbeAtMs < ANILIST_DURATION_RETRY_INTERVAL_MS) {
    return null;
  }
  anilistLastDurationProbeAtMs = now;

  try {
    const durationCandidate = await appState.mpvClient?.requestProperty("duration");
    const duration =
      typeof durationCandidate === "number" && Number.isFinite(durationCandidate)
        ? durationCandidate
        : null;
    if (duration && duration > 0 && anilistCurrentMediaKey === mediaKey) {
      anilistCurrentMediaDurationSec = duration;
      return duration;
    }
  } catch (error) {
    logger.warn("AniList duration probe failed:", error);
  }
  return null;
}

async function ensureAnilistMediaGuess(mediaKey: string): Promise<AnilistMediaGuess | null> {
  if (anilistCurrentMediaKey !== mediaKey) {
    return null;
  }
  if (anilistCurrentMediaGuess) {
    return anilistCurrentMediaGuess;
  }
  if (anilistCurrentMediaGuessPromise) {
    return anilistCurrentMediaGuessPromise;
  }

  const mediaPathForGuess = mediaRuntime.resolveMediaPathForJimaku(
    appState.currentMediaPath,
  );
  anilistCurrentMediaGuessPromise = guessAnilistMediaInfo(
    mediaPathForGuess,
    appState.currentMediaTitle,
  )
    .then((guess) => {
      if (anilistCurrentMediaKey === mediaKey) {
        anilistCurrentMediaGuess = guess;
      }
      return guess;
    })
    .finally(() => {
      if (anilistCurrentMediaKey === mediaKey) {
        anilistCurrentMediaGuessPromise = null;
      }
    });
  return anilistCurrentMediaGuessPromise;
}

function buildAnilistAttemptKey(mediaKey: string, episode: number): string {
  return `${mediaKey}::${episode}`;
}

function rememberAnilistAttemptedUpdateKey(key: string): void {
  anilistAttemptedUpdateKeys.add(key);
  if (anilistAttemptedUpdateKeys.size <= ANILIST_MAX_ATTEMPTED_UPDATE_KEYS) {
    return;
  }
  const oldestKey = anilistAttemptedUpdateKeys.values().next().value;
  if (typeof oldestKey === "string") {
    anilistAttemptedUpdateKeys.delete(oldestKey);
  }
}

async function processNextAnilistRetryUpdate(): Promise<{
  ok: boolean;
  message: string;
}> {
  const queued = anilistUpdateQueue.nextReady();
  refreshAnilistRetryQueueState();
  if (!queued) {
    return { ok: true, message: "AniList queue has no ready items." };
  }

  appState.anilistRetryQueueState.lastAttemptAt = Date.now();
  const accessToken = await refreshAnilistClientSecretState();
  if (!accessToken) {
    appState.anilistRetryQueueState.lastError = "AniList token unavailable for queued retry.";
    return { ok: false, message: appState.anilistRetryQueueState.lastError };
  }

  const result = await updateAnilistPostWatchProgress(
    accessToken,
    queued.title,
    queued.episode,
  );
  if (result.status === "updated" || result.status === "skipped") {
    anilistUpdateQueue.markSuccess(queued.key);
    rememberAnilistAttemptedUpdateKey(queued.key);
    appState.anilistRetryQueueState.lastError = null;
    refreshAnilistRetryQueueState();
    logger.info(`[AniList queue] ${result.message}`);
    return { ok: true, message: result.message };
  }

  anilistUpdateQueue.markFailure(queued.key, result.message);
  appState.anilistRetryQueueState.lastError = result.message;
  refreshAnilistRetryQueueState();
  return { ok: false, message: result.message };
}

async function maybeRunAnilistPostWatchUpdate(): Promise<void> {
  if (anilistUpdateInFlight) {
    return;
  }

  const resolved = getResolvedConfig();
  if (!isAnilistTrackingEnabled(resolved)) {
    return;
  }

  const mediaKey = getCurrentAnilistMediaKey();
  if (!mediaKey || !appState.mpvClient) {
    return;
  }
  if (anilistCurrentMediaKey !== mediaKey) {
    resetAnilistMediaTracking(mediaKey);
  }

  const watchedSeconds = appState.mpvClient.currentTimePos;
  if (
    !Number.isFinite(watchedSeconds) ||
    watchedSeconds < ANILIST_UPDATE_MIN_WATCH_SECONDS
  ) {
    return;
  }

  const duration = await maybeProbeAnilistDuration(mediaKey);
  if (!duration || duration <= 0) {
    return;
  }
  if (watchedSeconds / duration < ANILIST_UPDATE_MIN_WATCH_RATIO) {
    return;
  }

  const guess = await ensureAnilistMediaGuess(mediaKey);
  if (!guess?.title || !guess.episode || guess.episode <= 0) {
    return;
  }

  const attemptKey = buildAnilistAttemptKey(mediaKey, guess.episode);
  if (anilistAttemptedUpdateKeys.has(attemptKey)) {
    return;
  }

  anilistUpdateInFlight = true;
  try {
    await processNextAnilistRetryUpdate();

    const accessToken = await refreshAnilistClientSecretState();
    if (!accessToken) {
      anilistUpdateQueue.enqueue(attemptKey, guess.title, guess.episode);
      anilistUpdateQueue.markFailure(
        attemptKey,
        "cannot authenticate without anilist.accessToken",
      );
      refreshAnilistRetryQueueState();
      showMpvOsd("AniList: access token not configured");
      return;
    }
    const result = await updateAnilistPostWatchProgress(
      accessToken,
      guess.title,
      guess.episode,
    );
    if (result.status === "updated") {
      rememberAnilistAttemptedUpdateKey(attemptKey);
      anilistUpdateQueue.markSuccess(attemptKey);
      refreshAnilistRetryQueueState();
      showMpvOsd(result.message);
      logger.info(result.message);
      return;
    }
    if (result.status === "skipped") {
      rememberAnilistAttemptedUpdateKey(attemptKey);
      anilistUpdateQueue.markSuccess(attemptKey);
      refreshAnilistRetryQueueState();
      logger.info(result.message);
      return;
    }
    anilistUpdateQueue.enqueue(attemptKey, guess.title, guess.episode);
    anilistUpdateQueue.markFailure(attemptKey, result.message);
    refreshAnilistRetryQueueState();
    showMpvOsd(`AniList: ${result.message}`);
    logger.warn(result.message);
  } finally {
    anilistUpdateInFlight = false;
  }
}

function loadSubtitlePosition(): SubtitlePosition | null {
  appState.subtitlePosition = loadSubtitlePositionCore({
    currentMediaPath: appState.currentMediaPath,
    fallbackPosition: getResolvedConfig().subtitlePosition,
    subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
  });
  return appState.subtitlePosition;
}

function saveSubtitlePosition(position: SubtitlePosition): void {
  appState.subtitlePosition = position;
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
}

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
          appState.keybindings = resolveKeybindings(
            getResolvedConfig(),
            DEFAULT_KEYBINDINGS,
          );
        },
        createMpvClient: () => {
          appState.mpvClient = createMpvClientRuntimeService();
        },
        reloadConfig: () => {
          configService.reloadConfig();
          appLogger.logInfo(`Using config file: ${configService.getConfigPath()}`);
          void refreshAnilistClientSecretState({ force: true });
        },
        getResolvedConfig: () => getResolvedConfig(),
        getConfigWarnings: () => configService.getWarnings(),
        logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
        setLogLevel: (level: string, source: LogLevelSource) =>
          setLogLevel(level, source),
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
        defaultSecondarySubMode: "hover",
        defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
        hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
        startSubtitleWebsocket: (port: number) => {
          subtitleWsService.start(port, () => appState.currentSubText);
        },
        log: (message) => appLogger.logInfo(message),
        createMecabTokenizerAndCheck: async () => {
          const tokenizer = new MecabTokenizer();
          appState.mecabTokenizer = tokenizer;
          await tokenizer.checkAvailability();
        },
        createSubtitleTimingTracker: () => {
          const tracker = new SubtitleTimingTracker();
          appState.subtitleTimingTracker = tracker;
        },
        createImmersionTracker: () => {
          const config = getResolvedConfig();
          if (config.immersionTracking?.enabled === false) {
            logger.info("Immersion tracking disabled in config");
            return;
          }
          try {
            logger.debug(
              "Immersion tracker startup requested: creating tracker service.",
            );
            const dbPath = getConfiguredImmersionDbPath();
            logger.info(`Creating immersion tracker with dbPath=${dbPath}`);
            appState.immersionTracker = new ImmersionTrackerService({
              dbPath,
            });
            logger.debug("Immersion tracker initialized successfully.");
            if (appState.mpvClient && !appState.mpvClient.connected) {
              logger.info("Auto-connecting MPV client for immersion tracking");
              appState.mpvClient.connect();
            }
            seedImmersionTrackerFromCurrentMedia();
          } catch (error) {
            logger.warn("Immersion tracker startup failed; disabling tracking.", error);
            appState.immersionTracker = null;
          }
        },
        loadYomitanExtension: async () => {
          await loadYomitanExtension();
        },
        texthookerOnlyMode: appState.texthookerOnlyMode,
        shouldAutoInitializeOverlayRuntimeFromConfig: () =>
          shouldAutoInitializeOverlayRuntimeFromConfig(),
        initializeOverlayRuntime: () => initializeOverlayRuntime(),
        handleInitialArgs: () => handleInitialArgs(),
      }),
      onWillQuitCleanup: () => {
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
      },
      shouldRestoreWindowsOnActivate: () =>
        appState.overlayRuntimeInitialized && BrowserWindow.getAllWindows().length === 0,
      restoreWindowsOnActivate: () => {
        createMainWindow();
        createInvisibleWindow();
        overlayVisibilityRuntime.updateVisibleOverlayVisibility();
        overlayVisibilityRuntime.updateInvisibleOverlayVisibility();
      },
    }),
  }),
);

applyStartupState(appState, startupState);
void refreshAnilistClientSecretState({ force: true });
refreshAnilistRetryQueueState();

function handleCliCommand(
  args: CliArgs,
  source: CliCommandSource = "initial",
): void {
  handleCliCommandRuntimeServiceWithContext(args, source, {
    getSocketPath: () => appState.mpvSocketPath,
    setSocketPath: (socketPath: string) => {
      appState.mpvSocketPath = socketPath;
    },
    getClient: () => appState.mpvClient,
    showOsd: (text: string) => showMpvOsd(text),
    texthookerService,
    getTexthookerPort: () => appState.texthookerPort,
    setTexthookerPort: (port: number) => {
      appState.texthookerPort = port;
    },
    shouldOpenBrowser: () => getResolvedConfig().texthooker?.openBrowser !== false,
    openInBrowser: (url: string) => {
      void shell.openExternal(url).catch((error) => {
        logger.error(`Failed to open browser for texthooker URL: ${url}`, error);
      });
    },
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
    getAnilistStatus: () => getAnilistStatusSnapshot(),
    clearAnilistToken: () => clearAnilistTokenState(),
    openAnilistSetup: () => openAnilistSetupWindow(),
    getAnilistQueueStatus: () => getAnilistQueueStatusSnapshot(),
    retryAnilistQueueNow: () => processNextAnilistRetryUpdate(),
    openYomitanSettings: () => openYomitanSettings(),
    cycleSecondarySubMode: () => cycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
    stopApp: () => app.quit(),
    hasMainWindow: () => Boolean(overlayManager.getMainWindow()),
    getMultiCopyTimeoutMs: () => getConfiguredShortcuts().multiCopyTimeoutMs,
    schedule: (fn: () => void, delayMs: number) => setTimeout(fn, delayMs),
    log: (message: string) => {
      logger.info(message);
    },
    warn: (message: string) => {
      logger.warn(message);
    },
    error: (message: string, err: unknown) => {
      logger.error(message, err);
    },
  });
}

function handleInitialArgs(): void {
  if (!appState.initialArgs) return;
  if (
    !appState.texthookerOnlyMode &&
    appState.immersionTracker &&
    appState.mpvClient &&
    !appState.mpvClient.connected
  ) {
    logger.info("Auto-connecting MPV client for immersion tracking");
    appState.mpvClient.connect();
  }
  handleCliCommand(appState.initialArgs, "initial");
}

function bindMpvClientEventHandlers(mpvClient: MpvIpcClient): void {
  mpvClient.on("subtitle-change", ({ text }) => {
    appState.currentSubText = text;
    subtitleWsService.broadcast(text);
    void (async () => {
      if (getOverlayWindows().length > 0) {
        const subtitleData = await tokenizeSubtitle(text);
        broadcastToOverlayWindows("subtitle:set", subtitleData);
      }
    })();
  });
  mpvClient.on("subtitle-ass-change", ({ text }) => {
    appState.currentSubAssText = text;
    broadcastToOverlayWindows("subtitle-ass:set", text);
  });
  mpvClient.on("secondary-subtitle-change", ({ text }) => {
    broadcastToOverlayWindows("secondary-subtitle:set", text);
  });
  mpvClient.on("subtitle-timing", ({ text, start, end }) => {
    if (!text.trim()) {
      return;
    }
    appState.immersionTracker?.recordSubtitleLine(text, start, end);
    if (!appState.subtitleTimingTracker) {
      return;
    }
    appState.subtitleTimingTracker.recordSubtitle(text, start, end);
    void maybeRunAnilistPostWatchUpdate().catch((error) => {
      logger.error("AniList post-watch update failed unexpectedly", error);
    });
  });
  mpvClient.on("media-path-change", ({ path }) => {
    mediaRuntime.updateCurrentMediaPath(path);
    const mediaKey = getCurrentAnilistMediaKey();
    resetAnilistMediaTracking(mediaKey);
    if (mediaKey) {
      void maybeProbeAnilistDuration(mediaKey);
      void ensureAnilistMediaGuess(mediaKey);
    }
    syncImmersionTrackerFromCurrentMediaState();
  });
  mpvClient.on("media-title-change", ({ title }) => {
    mediaRuntime.updateCurrentMediaTitle(title);
    anilistCurrentMediaGuess = null;
    anilistCurrentMediaGuessPromise = null;
    appState.immersionTracker?.handleMediaTitleUpdate(title);
    syncImmersionTrackerFromCurrentMediaState();
  });
  mpvClient.on("time-pos-change", ({ time }) => {
    appState.immersionTracker?.recordPlaybackPosition(time);
  });
  mpvClient.on("pause-change", ({ paused }) => {
    appState.immersionTracker?.recordPauseState(paused);
  });
  mpvClient.on("subtitle-metrics-change", ({ patch }) => {
    updateMpvSubtitleRenderMetrics(patch);
  });
  mpvClient.on("secondary-subtitle-visibility", ({ visible }) => {
    appState.previousSecondarySubVisibility = visible;
  });
}

function createMpvClientRuntimeService(): MpvIpcClient {
  const mpvClient = new MpvIpcClient(appState.mpvSocketPath, {
    getResolvedConfig: () => getResolvedConfig(),
    autoStartOverlay: appState.autoStartOverlay,
    setOverlayVisible: (visible: boolean) => setOverlayVisible(visible),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      shouldBindVisibleOverlayToMpvSubVisibility(),
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getReconnectTimer: () => appState.reconnectTimer,
    setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
      appState.reconnectTimer = timer;
    },
  });
  bindMpvClientEventHandlers(mpvClient);
  mpvClient.connect();
  return mpvClient;
}

function updateMpvSubtitleRenderMetrics(
  patch: Partial<MpvSubtitleRenderMetrics>,
): void {
  const { next, changed } = applyMpvSubtitleRenderMetricsPatch(
    appState.mpvSubtitleRenderMetrics,
    patch,
  );
  if (!changed) return;
  appState.mpvSubtitleRenderMetrics = next;
  broadcastToOverlayWindows(
    "mpv-subtitle-render-metrics:set",
    appState.mpvSubtitleRenderMetrics,
  );
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
      getJlptEnabled: () =>
        getResolvedConfig().subtitleStyle.enableJlpt,
      getFrequencyDictionaryEnabled: () =>
        getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
      getFrequencyRank: (text) => appState.frequencyRankLookup(text),
      getYomitanGroupDebugEnabled: () => appState.overlayDebugVisualizationEnabled,
      getMecabTokenizer: () => appState.mecabTokenizer,
    }),
  );
}

function updateVisibleOverlayBounds(geometry: WindowGeometry): void {
  overlayManager.setOverlayWindowBounds("visible", geometry);
}

function updateInvisibleOverlayBounds(geometry: WindowGeometry): void {
  overlayManager.setOverlayWindowBounds("invisible", geometry);
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

function createOverlayWindow(kind: "visible" | "invisible"): BrowserWindow {
  return createOverlayWindowCore(
    kind,
    {
      isDev,
      overlayDebugVisualizationEnabled: appState.overlayDebugVisualizationEnabled,
      ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
      onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
      setOverlayDebugVisualizationEnabled: (enabled) =>
        setOverlayDebugVisualizationEnabled(enabled),
      isOverlayVisible: (windowKind) =>
        windowKind === "visible"
          ? overlayManager.getVisibleOverlayVisible()
          : overlayManager.getInvisibleOverlayVisible(),
      tryHandleOverlayShortcutLocalFallback: (input) =>
        overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(input),
      onWindowClosed: (windowKind) => {
        if (windowKind === "visible") {
          overlayManager.setMainWindow(null);
        } else {
          overlayManager.setInvisibleWindow(null);
        }
      },
    },
  );
}

function createMainWindow(): BrowserWindow {
  const window = createOverlayWindow("visible");
  overlayManager.setMainWindow(window);
  return window;
}
function createInvisibleWindow(): BrowserWindow {
  const window = createOverlayWindow("invisible");
  overlayManager.setInvisibleWindow(window);
  return window;
}

function initializeOverlayRuntime(): void {
  if (appState.overlayRuntimeInitialized) {
    return;
  }
  const result = initializeOverlayRuntimeCore(
    {
      backendOverride: appState.backendOverride,
      getInitialInvisibleOverlayVisibility: () =>
        getInitialInvisibleOverlayVisibility(),
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
      isInvisibleOverlayVisible: () =>
        overlayManager.getInvisibleOverlayVisible(),
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
      getKnownWordCacheStatePath: () =>
        path.join(USER_DATA_PATH, "known-words-cache.json"),
    },
  );
  overlayManager.setInvisibleOverlayVisible(result.invisibleOverlayVisible);
  appState.overlayRuntimeInitialized = true;
}

function openYomitanSettings(): void {
  openYomitanSettingsWindow(
    {
      yomitanExt: appState.yomitanExt,
      getExistingWindow: () => appState.yomitanSettingsWindow,
      setWindow: (window: BrowserWindow | null) => {
        appState.yomitanSettingsWindow = window;
      },
    },
  );
}
function registerGlobalShortcuts(): void {
  registerGlobalShortcutsCore(
    {
      shortcuts: getConfiguredShortcuts(),
      onToggleVisibleOverlay: () => toggleVisibleOverlay(),
      onToggleInvisibleOverlay: () => toggleInvisibleOverlay(),
      onOpenYomitanSettings: () => openYomitanSettings(),
      isDev,
      getMainWindow: () => overlayManager.getMainWindow(),
    },
  );
}

function getConfiguredShortcuts() { return resolveConfiguredShortcuts(getResolvedConfig(), DEFAULT_CONFIG); }

function cycleSecondarySubMode(): void {
  cycleSecondarySubModeCore(
    {
      getSecondarySubMode: () => appState.secondarySubMode,
      setSecondarySubMode: (mode: SecondarySubMode) => {
        appState.secondarySubMode = mode;
      },
      getLastSecondarySubToggleAtMs: () => appState.lastSecondarySubToggleAtMs,
      setLastSecondarySubToggleAtMs: (timestampMs: number) => {
        appState.lastSecondarySubToggleAtMs = timestampMs;
      },
      broadcastSecondarySubMode: (mode: SecondarySubMode) => {
        broadcastToOverlayWindows("secondary-subtitle:mode", mode);
      },
      showMpvOsd: (text: string) => showMpvOsd(text),
    },
  );
}

function showMpvOsd(text: string): void {
  appendToMpvLog(`[OSD] ${text}`);
  showMpvOsdRuntime(
    appState.mpvClient,
    text,
    (line) => {
      logger.info(line);
    },
  );
}

function appendToMpvLog(message: string): void {
  try {
    fs.mkdirSync(path.dirname(DEFAULT_MPV_LOG_PATH), { recursive: true });
    fs.appendFileSync(
      DEFAULT_MPV_LOG_PATH,
      `[${new Date().toISOString()}] ${message}\n`,
      { encoding: "utf8" },
    );
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

function getSubsyncRuntimeServiceParams() {
  return createSubsyncRuntimeServiceInputFromState({
    getMpvClient: () => appState.mpvClient,
    getResolvedSubsyncConfig: () => getSubsyncConfig(getResolvedConfig().subsync),
    getSubsyncInProgress: () => appState.subsyncInProgress,
    setSubsyncInProgress: (inProgress: boolean) => {
      appState.subsyncInProgress = inProgress;
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
    openManualPicker: (payload: SubsyncManualPayload) => {
      sendToActiveOverlayWindow("subsync:open-manual", payload, {
        restoreOnModalClose: "subsync",
      });
    },
  });
}

async function triggerSubsyncFromConfig(): Promise<void> {
  await triggerSubsyncFromConfigRuntime(getSubsyncRuntimeServiceParams());
}

function cancelPendingMultiCopy(): void {
  multiCopySession.cancel();
}

function startPendingMultiCopy(timeoutMs: number): void {
  multiCopySession.start({
    timeoutMs,
    onDigit: (count) => handleMultiCopyDigit(count),
    messages: {
      prompt: "Copy how many lines? Press 1-9 (Esc to cancel)",
      timeout: "Copy timeout",
      cancelled: "Cancelled",
    },
  });
}

function handleMultiCopyDigit(count: number): void {
  handleMultiCopyDigitCore(
    count,
    {
      subtitleTimingTracker: appState.subtitleTimingTracker,
      writeClipboardText: (text) => clipboard.writeText(text),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

function copyCurrentSubtitle(): void {
  copyCurrentSubtitleCore(
    {
      subtitleTimingTracker: appState.subtitleTimingTracker,
      writeClipboardText: (text) => clipboard.writeText(text),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function updateLastCardFromClipboard(): Promise<void> {
  await updateLastCardFromClipboardCore(
    {
      ankiIntegration: appState.ankiIntegration,
      readClipboardText: () => clipboard.readText(),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function refreshKnownWordCache(): Promise<void> {
  if (!appState.ankiIntegration) {
    throw new Error("AnkiConnect integration not enabled");
  }

  await appState.ankiIntegration.refreshKnownWordCache();
}

async function triggerFieldGrouping(): Promise<void> {
  await triggerFieldGroupingCore(
    {
      ankiIntegration: appState.ankiIntegration,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function markLastCardAsAudioCard(): Promise<void> {
  await markLastCardAsAudioCardCore(
    {
      ankiIntegration: appState.ankiIntegration,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function mineSentenceCard(): Promise<void> {
  const created = await mineSentenceCardCore(
    {
      ankiIntegration: appState.ankiIntegration,
      mpvClient: appState.mpvClient,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
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
      prompt: "Mine how many lines? Press 1-9 (Esc to cancel)",
      timeout: "Mine sentence timeout",
      cancelled: "Cancelled",
    },
  });
}

function handleMineSentenceDigit(count: number): void {
  handleMineSentenceDigitCore(
    count,
    {
      subtitleTimingTracker: appState.subtitleTimingTracker,
      ankiIntegration: appState.ankiIntegration,
      getCurrentSecondarySubText: () =>
        appState.mpvClient?.currentSecondarySubText || undefined,
      showMpvOsd: (text) => showMpvOsd(text),
      logError: (message, err) => {
        logger.error(message, err);
      },
      onCardsMined: (cards) => {
        appState.immersionTracker?.recordCardsMined(cards);
      },
    },
  );
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
    updateVisibleOverlayVisibility: () =>
      overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () =>
      overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      overlayVisibilityRuntime.syncInvisibleOverlayMousePassthrough(),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      shouldBindVisibleOverlayToMpvSubVisibility(),
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
function setOverlayVisible(visible: boolean): void { setVisibleOverlayVisible(visible); }
function toggleOverlay(): void { toggleVisibleOverlay(); }
function handleOverlayModalClosed(modal: OverlayHostedModal): void {
  overlayModalRuntime.handleOverlayModalClosed(modal);
}

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcRuntime(command, {
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    cycleRuntimeOption: (id, direction) => {
      if (!appState.runtimeOptionsManager) {
        return { ok: false, error: "Runtime options manager unavailable" };
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

async function runSubsyncManualFromIpc(
  request: SubsyncManualRunRequest,
): Promise<SubsyncResult> {
  return runSubsyncManualFromIpcRuntime(request, getSubsyncRuntimeServiceParams());
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
    getCurrentSubtitleAss: () => appState.currentSubAssText,
    getMpvSubtitleRenderMetrics: () => appState.mpvSubtitleRenderMetrics,
    getSubtitlePosition: () => loadSubtitlePosition(),
      getSubtitleStyle: () => {
        const resolvedConfig = getResolvedConfig();
        if (!resolvedConfig.subtitleStyle) {
          return null;
        }

        return {
          ...resolvedConfig.subtitleStyle,
          nPlusOneColor: resolvedConfig.ankiConnect.nPlusOne.nPlusOne,
          knownWordColor: resolvedConfig.ankiConnect.nPlusOne.knownWord,
          enableJlpt: resolvedConfig.subtitleStyle.enableJlpt,
          frequencyDictionary:
            resolvedConfig.subtitleStyle.frequencyDictionary,
        };
      },
    saveSubtitlePosition: (position: unknown) =>
      saveSubtitlePosition(position as SubtitlePosition),
    getMecabTokenizer: () => appState.mecabTokenizer,
    handleMpvCommand: (command: (string | number)[]) =>
      handleMpvCommandFromIpc(command),
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
    getAnilistStatus: () => getAnilistStatusSnapshot(),
    clearAnilistToken: () => clearAnilistTokenState(),
    openAnilistSetup: () => openAnilistSetupWindow(),
    getAnilistQueueStatus: () => getAnilistQueueStatusSnapshot(),
    retryAnilistQueueNow: () => processNextAnilistRetryUpdate(),
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
    getKnownWordCacheStatePath: () =>
      path.join(USER_DATA_PATH, "known-words-cache.json"),
    showDesktopNotification,
    createFieldGroupingCallback: () => createFieldGroupingCallback(),
    broadcastRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    getFieldGroupingResolver: () => getFieldGroupingResolver(),
    setFieldGroupingResolver: (
      resolver: ((choice: KikuFieldGroupingChoice) => void) | null,
    ) => setFieldGroupingResolver(resolver),
    parseMediaInfo: (mediaPath: string | null) =>
      parseMediaInfo(mediaRuntime.resolveMediaPathForJimaku(mediaPath)),
    getCurrentMediaPath: () => appState.currentMediaPath,
    jimakuFetchJson: <T>(
      endpoint: string,
      query?: Record<string, string | number | boolean | null | undefined>,
    ): Promise<JimakuApiResponse<T>> => jimakuFetchJson<T>(endpoint, query),
    getJimakuMaxEntryResults: () => getJimakuMaxEntryResults(),
    getJimakuLanguagePreference: () => getJimakuLanguagePreference(),
    resolveJimakuApiKey: () => resolveJimakuApiKey(),
    isRemoteMediaPath: (mediaPath: string) => isRemoteMediaPath(mediaPath),
    downloadToFile: (
      url: string,
      destPath: string,
      headers: Record<string, string>,
    ) => downloadToFile(url, destPath, headers),
  }),
});
