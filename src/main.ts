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
import { BaseWindowTracker } from "./window-trackers";
import type {
  JimakuApiResponse,
  JimakuLanguagePreference,
  SubtitleData,
  SubtitlePosition,
  Keybinding,
  WindowGeometry,
  SecondarySubMode,
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncResult,
  KikuFieldGroupingChoice,
  KikuMergePreviewRequest,
  KikuMergePreviewResponse,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
  MpvSubtitleRenderMetrics,
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
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
  MpvIpcClient,
  SubtitleWebSocketService,
  TexthookerService,
  applyMpvSubtitleRenderMetricsPatchService,
  broadcastRuntimeOptionsChangedRuntimeService,
  copyCurrentSubtitleService,
  createAppLifecycleDepsRuntimeService,
  createCliCommandDepsRuntimeService,
  createOverlayManagerService,
  createFieldGroupingOverlayRuntimeService,
  createIpcDepsRuntimeService,
  createNumericShortcutRuntimeService,
  createOverlayContentMeasurementStoreService,
  createOverlayShortcutRuntimeHandlers,
  createOverlayWindowService,
  createTokenizerDepsRuntimeService,
  cycleSecondarySubModeService,
  enforceOverlayLayerOrderService,
  ensureOverlayWindowLevelService,
  getInitialInvisibleOverlayVisibilityService,
  getJimakuLanguagePreferenceService,
  getJimakuMaxEntryResultsService,
  handleCliCommandService,
  handleMineSentenceDigitService,
  handleMpvCommandFromIpcService,
  handleMultiCopyDigitService,
  hasMpvWebsocketPlugin,
  initializeOverlayRuntimeService,
  isAutoUpdateEnabledRuntimeService,
  jimakuFetchJsonService,
  loadSubtitlePositionService,
  loadYomitanExtensionService,
  markLastCardAsAudioCardService,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
  mineSentenceCardService,
  openYomitanSettingsWindow,
  playNextSubtitleRuntimeService,
  refreshOverlayShortcutsRuntimeService,
  registerAnkiJimakuIpcRuntimeService,
  registerGlobalShortcutsService,
  registerIpcHandlersService,
  registerOverlayShortcutsService,
  replayCurrentSubtitleRuntimeService,
  resolveJimakuApiKeyService,
  runStartupBootstrapRuntimeService,
  runSubsyncManualFromIpcRuntimeService,
  saveSubtitlePositionService,
  sendMpvCommandRuntimeService,
  setInvisibleOverlayVisibleService,
  setMpvSubVisibilityRuntimeService,
  setOverlayDebugVisualizationEnabledRuntimeService,
  setVisibleOverlayVisibleService,
  shouldAutoInitializeOverlayRuntimeFromConfigService,
  shouldBindVisibleOverlayToMpvSubVisibilityService,
  shortcutMatchesInputForLocalFallback,
  showMpvOsdRuntimeService,
  startAppLifecycleService,
  syncInvisibleOverlayMousePassthroughService,
  syncOverlayShortcutsRuntimeService,
  tokenizeSubtitleService,
  triggerFieldGroupingService,
  triggerSubsyncFromConfigRuntimeService,
  unregisterOverlayShortcutsRuntimeService,
  updateCurrentMediaPathService,
  updateInvisibleOverlayVisibilityService,
  updateLastCardFromClipboardService,
  updateVisibleOverlayVisibilityService,
} from "./core/services";
import { runOverlayShortcutLocalFallback } from "./core/services/overlay-shortcut-handler";
import { runAppReadyRuntimeService } from "./core/services/startup-service";
import {
  applyRuntimeOptionResultRuntimeService,
  cycleRuntimeOptionFromIpcRuntimeService,
  setRuntimeOptionFromIpcRuntimeService,
} from "./core/services/runtime-options-ipc-service";
import {
  ConfigService,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
  SPECIAL_COMMANDS,
} from "./config";

if (process.platform === "linux") {
  app.commandLine.appendSwitch("enable-features", "GlobalShortcutsPortal");
}

const DEFAULT_TEXTHOOKER_PORT = 5174;
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
const configService = new ConfigService(CONFIG_DIR);
const isDev =
  process.argv.includes("--dev") || process.argv.includes("--debug");
const texthookerService = new TexthookerService();
const subtitleWsService = new SubtitleWebSocketService();
const appLogger = {
  logInfo: (message: string) => {
    console.log(message);
  },
  logWarning: (message: string) => {
    console.warn(message);
  },
  logNoRunningInstance: () => {
    console.error("No running instance. Use --start to launch the app.");
  },
  logConfigWarning: (warning: {
    path: string;
    message: string;
    value: unknown;
    fallback: unknown;
  }) => {
    console.warn(
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

let yomitanExt: Extension | null = null;
let yomitanSettingsWindow: BrowserWindow | null = null;
let yomitanParserWindow: BrowserWindow | null = null;
let yomitanParserReadyPromise: Promise<void> | null = null;
let yomitanParserInitPromise: Promise<boolean> | null = null;
let mpvClient: MpvIpcClient | null = null;
let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
let currentSubText = "";
let currentSubAssText = "";
let windowTracker: BaseWindowTracker | null = null;
let subtitlePosition: SubtitlePosition | null = null;
let currentMediaPath: string | null = null;
let pendingSubtitlePosition: SubtitlePosition | null = null;
let mecabTokenizer: MecabTokenizer | null = null;
let keybindings: Keybinding[] = [];
let subtitleTimingTracker: SubtitleTimingTracker | null = null;
let ankiIntegration: AnkiIntegration | null = null;
let secondarySubMode: SecondarySubMode = "hover";
let lastSecondarySubToggleAtMs = 0;
let previousSecondarySubVisibility: boolean | null = null;
let mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics = {
  ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
};

let shortcutsRegistered = false;
let overlayRuntimeInitialized = false;
let fieldGroupingResolver: ((choice: KikuFieldGroupingChoice) => void) | null =
  null;
let fieldGroupingResolverSequence = 0;
let runtimeOptionsManager: RuntimeOptionsManager | null = null;
let trackerNotReadyWarningShown = false;
let overlayDebugVisualizationEnabled = false;
const overlayManager = createOverlayManagerService();
const overlayContentMeasurementStore = createOverlayContentMeasurementStoreService({
  now: () => Date.now(),
  warn: (message: string) => {
    console.warn(message);
  },
});
type OverlayHostedModal = "runtime-options" | "subsync";
const restoreVisibleOverlayOnModalClose = new Set<OverlayHostedModal>();

function getFieldGroupingResolver(): ((choice: KikuFieldGroupingChoice) => void) | null {
  return fieldGroupingResolver;
}

function setFieldGroupingResolver(
  resolver: ((choice: KikuFieldGroupingChoice) => void) | null,
): void {
  if (!resolver) {
    fieldGroupingResolver = null;
    return;
  }
  const sequence = ++fieldGroupingResolverSequence;
  const wrappedResolver = (choice: KikuFieldGroupingChoice): void => {
    if (sequence !== fieldGroupingResolverSequence) return;
    resolver(choice);
  };
  fieldGroupingResolver = wrappedResolver;
}

const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntimeService<OverlayHostedModal>({
  getMainWindow: () => overlayManager.getMainWindow(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
  setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  setInvisibleOverlayVisible: (visible) => setInvisibleOverlayVisible(visible),
  getResolver: () => getFieldGroupingResolver(),
  setResolver: (resolver) => setFieldGroupingResolver(resolver),
  getRestoreVisibleOverlayOnModalClose: () => restoreVisibleOverlayOnModalClose,
});
const sendToVisibleOverlay = fieldGroupingOverlayRuntime.sendToVisibleOverlay;
const createFieldGroupingCallback =
  fieldGroupingOverlayRuntime.createFieldGroupingCallback;

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, "subtitle-positions");

function getRuntimeOptionsState(): RuntimeOptionState[] { if (!runtimeOptionsManager) return []; return runtimeOptionsManager.listOptions(); }

function getOverlayWindows(): BrowserWindow[] {
  return overlayManager.getOverlayWindows();
}

function broadcastToOverlayWindows(channel: string, ...args: unknown[]): void {
  overlayManager.broadcastToOverlayWindows(channel, ...args);
}

function broadcastRuntimeOptionsChanged(): void {
  broadcastRuntimeOptionsChangedRuntimeService(
    () => getRuntimeOptionsState(),
    (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  );
}

function setOverlayDebugVisualizationEnabled(enabled: boolean): void {
  setOverlayDebugVisualizationEnabledRuntimeService(
    overlayDebugVisualizationEnabled,
    enabled,
    (next) => {
      overlayDebugVisualizationEnabled = next;
    },
    (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  );
}

function openRuntimeOptionsPalette(): void { sendToVisibleOverlay("runtime-options:open", undefined, { restoreOnModalClose: "runtime-options" }); }

function getResolvedConfig() { return configService.getConfig(); }

function getInitialInvisibleOverlayVisibility(): boolean {
  return getInitialInvisibleOverlayVisibilityService(
    getResolvedConfig(),
    process.platform,
  );
}

function shouldAutoInitializeOverlayRuntimeFromConfig(): boolean {
  return shouldAutoInitializeOverlayRuntimeFromConfigService(getResolvedConfig());
}

function shouldBindVisibleOverlayToMpvSubVisibility(): boolean {
  return shouldBindVisibleOverlayToMpvSubVisibilityService(getResolvedConfig());
}

function isAutoUpdateEnabledRuntime(): boolean {
  return isAutoUpdateEnabledRuntimeService(
    getResolvedConfig(),
    runtimeOptionsManager,
  );
}

function getJimakuLanguagePreference(): JimakuLanguagePreference { return getJimakuLanguagePreferenceService(() => getResolvedConfig(), DEFAULT_CONFIG.jimaku.languagePreference); }

function getJimakuMaxEntryResults(): number { return getJimakuMaxEntryResultsService(() => getResolvedConfig(), DEFAULT_CONFIG.jimaku.maxEntryResults); }

async function resolveJimakuApiKey(): Promise<string | null> { return resolveJimakuApiKeyService(() => getResolvedConfig()); }

async function jimakuFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined> = {},
): Promise<JimakuApiResponse<T>> {
  return jimakuFetchJsonService<T>(endpoint, query, {
    getResolvedConfig: () => getResolvedConfig(),
    defaultBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
    defaultMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
    defaultLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  });
}

function loadSubtitlePosition(): SubtitlePosition | null {
  subtitlePosition = loadSubtitlePositionService({
    currentMediaPath,
    fallbackPosition: getResolvedConfig().subtitlePosition,
    subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
  });
  return subtitlePosition;
}

function saveSubtitlePosition(position: SubtitlePosition): void {
  subtitlePosition = position;
  saveSubtitlePositionService({
    position,
    currentMediaPath,
    subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
    onQueuePending: (queued) => {
      pendingSubtitlePosition = queued;
    },
    onPersisted: () => {
      pendingSubtitlePosition = null;
    },
  });
}

function updateCurrentMediaPath(mediaPath: unknown): void {
  updateCurrentMediaPathService({
    mediaPath,
    currentMediaPath,
    pendingSubtitlePosition,
    subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
    loadSubtitlePosition: () => loadSubtitlePosition(),
    setCurrentMediaPath: (nextPath) => {
      currentMediaPath = nextPath;
    },
    clearPendingSubtitlePosition: () => {
      pendingSubtitlePosition = null;
    },
    setSubtitlePosition: (position) => {
      subtitlePosition = position;
    },
    broadcastSubtitlePosition: (position) => {
      broadcastToOverlayWindows("subtitle-position:set", position);
    },
  });
}

let subsyncInProgress = false;
let initialArgs: CliArgs;
let mpvSocketPath = getDefaultSocketPath();
let texthookerPort = DEFAULT_TEXTHOOKER_PORT;
let backendOverride: string | null = null;
let autoStartOverlay = false;
let texthookerOnlyMode = false;

const startupState = runStartupBootstrapRuntimeService({
  argv: process.argv,
  parseArgs: (argv) => parseArgs(argv),
  setLogLevelEnv: (level) => {
    process.env.SUBMINER_LOG_LEVEL = level;
  },
  enableVerboseLogging: () => {
    process.env.SUBMINER_LOG_LEVEL = "debug";
  },
  forceX11Backend: (args) => {
    forceX11Backend(args);
  },
  enforceUnsupportedWaylandMode: (args) => {
    enforceUnsupportedWaylandMode(args);
  },
  getDefaultSocketPath: () => getDefaultSocketPath(),
  defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
  runGenerateConfigFlow: (args) => {
    if (!args.generateConfig || shouldStartApp(args)) {
      return false;
    }
    generateDefaultConfigFile(args, {
      configDir: CONFIG_DIR,
      defaultConfig: DEFAULT_CONFIG,
      generateTemplate: (config) => generateConfigTemplate(config as never),
    })
      .then((exitCode) => {
        process.exitCode = exitCode;
        app.quit();
      })
      .catch((error: Error) => {
        console.error(`Failed to generate config: ${error.message}`);
        process.exitCode = 1;
        app.quit();
      });
    return true;
  },
  startAppLifecycle: (args) => {
    startAppLifecycleService(args, createAppLifecycleDepsRuntimeService({
      app,
      platform: process.platform,
      shouldStartApp: (nextArgs) => shouldStartApp(nextArgs),
      parseArgs: (argv) => parseArgs(argv),
      handleCliCommand: (nextArgs, source) => handleCliCommand(nextArgs, source),
      printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
      logNoRunningInstance: () => appLogger.logNoRunningInstance(),
      onReady: async () => {
        await runAppReadyRuntimeService({
          loadSubtitlePosition: () => loadSubtitlePosition(),
          resolveKeybindings: () => {
            keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);
          },
          createMpvClient: () => {
            mpvClient = new MpvIpcClient(
              mpvSocketPath,
              {
                getResolvedConfig: () => getResolvedConfig(),
                autoStartOverlay,
                setOverlayVisible: (visible) => setOverlayVisible(visible),
                shouldBindVisibleOverlayToMpvSubVisibility: () =>
                  shouldBindVisibleOverlayToMpvSubVisibility(),
                isVisibleOverlayVisible: () =>
                  overlayManager.getVisibleOverlayVisible(),
                getReconnectTimer: () => reconnectTimer,
                setReconnectTimer: (timer) => {
                  reconnectTimer = timer;
                },
                getCurrentSubText: () => currentSubText,
                setCurrentSubText: (text) => {
                  currentSubText = text;
                },
                setCurrentSubAssText: (text) => {
                  currentSubAssText = text;
                },
                getSubtitleTimingTracker: () => subtitleTimingTracker,
                subtitleWsBroadcast: (text) => {
                  subtitleWsService.broadcast(text);
                },
                getOverlayWindowsCount: () => getOverlayWindows().length,
                tokenizeSubtitle: (text) => tokenizeSubtitle(text),
                broadcastToOverlayWindows: (channel, ...channelArgs) => {
                  broadcastToOverlayWindows(channel, ...channelArgs);
                },
                updateCurrentMediaPath: (mediaPath) => {
                  updateCurrentMediaPath(mediaPath);
                },
                updateMpvSubtitleRenderMetrics: (patch) => {
                  updateMpvSubtitleRenderMetrics(patch);
                },
                getMpvSubtitleRenderMetrics: () => mpvSubtitleRenderMetrics,
                setPreviousSecondarySubVisibility: (value) => {
                  previousSecondarySubVisibility = value;
                },
                showMpvOsd: (text) => {
                  showMpvOsd(text);
                },
              },
            );
          },
          reloadConfig: () => {
            configService.reloadConfig();
            appLogger.logInfo(`Using config file: ${configService.getConfigPath()}`);
          },
          getResolvedConfig: () => getResolvedConfig(),
          getConfigWarnings: () => configService.getWarnings(),
          logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
          initRuntimeOptionsManager: () => {
            runtimeOptionsManager = new RuntimeOptionsManager(
              () => configService.getConfig().ankiConnect,
              {
                applyAnkiPatch: (patch) => {
                  if (ankiIntegration) {
                    ankiIntegration.applyRuntimeConfigPatch(patch);
                  }
                },
                onOptionsChanged: () => {
                  broadcastRuntimeOptionsChanged();
                  refreshOverlayShortcuts();
                },
              },
            );
          },
          setSecondarySubMode: (mode) => {
            secondarySubMode = mode;
          },
          defaultSecondarySubMode: "hover",
          defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
          hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
          startSubtitleWebsocket: (port) => {
            subtitleWsService.start(port, () => currentSubText);
          },
          log: (message) => appLogger.logInfo(message),
          createMecabTokenizerAndCheck: async () => {
            const tokenizer = new MecabTokenizer();
            mecabTokenizer = tokenizer;
            await tokenizer.checkAvailability();
          },
          createSubtitleTimingTracker: () => {
            const tracker = new SubtitleTimingTracker();
            subtitleTimingTracker = tracker;
          },
          loadYomitanExtension: async () => {
            await loadYomitanExtension();
          },
          texthookerOnlyMode,
          shouldAutoInitializeOverlayRuntimeFromConfig: () =>
            shouldAutoInitializeOverlayRuntimeFromConfig(),
          initializeOverlayRuntime: () => initializeOverlayRuntime(),
          handleInitialArgs: () => handleInitialArgs(),
        });
      },
      onWillQuitCleanup: () => {
        globalShortcut.unregisterAll();
        subtitleWsService.stop();
        texthookerService.stop();
        if (yomitanParserWindow && !yomitanParserWindow.isDestroyed()) {
          yomitanParserWindow.destroy();
        }
        yomitanParserWindow = null;
        yomitanParserReadyPromise = null;
        yomitanParserInitPromise = null;
        if (windowTracker) {
          windowTracker.stop();
        }
        if (mpvClient && mpvClient.socket) {
          mpvClient.socket.destroy();
        }
        if (reconnectTimer) {
          clearTimeout(reconnectTimer);
        }
        if (subtitleTimingTracker) {
          subtitleTimingTracker.destroy();
        }
        if (ankiIntegration) {
          ankiIntegration.destroy();
        }
      },
      shouldRestoreWindowsOnActivate: () =>
        overlayRuntimeInitialized && BrowserWindow.getAllWindows().length === 0,
      restoreWindowsOnActivate: () => {
        createMainWindow();
        createInvisibleWindow();
        updateVisibleOverlayVisibility();
        updateInvisibleOverlayVisibility();
      },
    }));
  },
});

initialArgs = startupState.initialArgs;
mpvSocketPath = startupState.mpvSocketPath;
texthookerPort = startupState.texthookerPort;
backendOverride = startupState.backendOverride;
autoStartOverlay = startupState.autoStartOverlay;
texthookerOnlyMode = startupState.texthookerOnlyMode;

function handleCliCommand(
  args: CliArgs,
  source: CliCommandSource = "initial",
): void {
  const deps = createCliCommandDepsRuntimeService({
    mpv: {
      getSocketPath: () => mpvSocketPath,
      setSocketPath: (socketPath) => {
        mpvSocketPath = socketPath;
      },
      getClient: () => mpvClient,
      showOsd: (text) => showMpvOsd(text),
    },
    texthooker: {
      service: texthookerService,
      getPort: () => texthookerPort,
      setPort: (port) => {
        texthookerPort = port;
      },
      shouldOpenBrowser: () => getResolvedConfig().texthooker?.openBrowser !== false,
      openInBrowser: (url) => {
        void shell.openExternal(url).catch((error) => {
          console.error(`Failed to open browser for texthooker URL: ${url}`, error);
        });
      },
    },
    overlay: {
      isInitialized: () => overlayRuntimeInitialized,
      initialize: () => initializeOverlayRuntime(),
      toggleVisible: () => toggleVisibleOverlay(),
      toggleInvisible: () => toggleInvisibleOverlay(),
      setVisible: (visible) => setVisibleOverlayVisible(visible),
      setInvisible: (visible) => setInvisibleOverlayVisible(visible),
    },
    mining: {
      copyCurrentSubtitle: () => copyCurrentSubtitle(),
      startPendingMultiCopy: (timeoutMs) => startPendingMultiCopy(timeoutMs),
      mineSentenceCard: () => mineSentenceCard(),
      startPendingMineSentenceMultiple: (timeoutMs) =>
        startPendingMineSentenceMultiple(timeoutMs),
      updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
      triggerFieldGrouping: () => triggerFieldGrouping(),
      triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
      markLastCardAsAudioCard: () => markLastCardAsAudioCard(),
    },
    ui: {
      openYomitanSettings: () => openYomitanSettings(),
      cycleSecondarySubMode: () => cycleSecondarySubMode(),
      openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
      printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
    },
    app: {
      stop: () => app.quit(),
      hasMainWindow: () => Boolean(overlayManager.getMainWindow()),
    },
    getMultiCopyTimeoutMs: () => getConfiguredShortcuts().multiCopyTimeoutMs,
    schedule: (fn, delayMs) => setTimeout(fn, delayMs),
    log: (message) => {
      console.log(message);
    },
    warn: (message) => {
      console.warn(message);
    },
    error: (message, err) => {
      console.error(message, err);
    },
  });
  handleCliCommandService(args, source, deps);
}

function handleInitialArgs(): void {
  handleCliCommand(initialArgs, "initial");
}

function updateMpvSubtitleRenderMetrics(
  patch: Partial<MpvSubtitleRenderMetrics>,
): void {
  const { next, changed } = applyMpvSubtitleRenderMetricsPatchService(
    mpvSubtitleRenderMetrics,
    patch,
  );
  if (!changed) return;
  mpvSubtitleRenderMetrics = next;
  broadcastToOverlayWindows(
    "mpv-subtitle-render-metrics:set",
    mpvSubtitleRenderMetrics,
  );
}

async function tokenizeSubtitle(text: string): Promise<SubtitleData> {
  return tokenizeSubtitleService(
    text,
    createTokenizerDepsRuntimeService({
      getYomitanExt: () => yomitanExt,
      getYomitanParserWindow: () => yomitanParserWindow,
      setYomitanParserWindow: (window) => {
        yomitanParserWindow = window;
      },
      getYomitanParserReadyPromise: () => yomitanParserReadyPromise,
      setYomitanParserReadyPromise: (promise) => {
        yomitanParserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () => yomitanParserInitPromise,
      setYomitanParserInitPromise: (promise) => {
        yomitanParserInitPromise = promise;
      },
      getMecabTokenizer: () => mecabTokenizer,
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
  ensureOverlayWindowLevelService(window);
}

function enforceOverlayLayerOrder(): void {
  enforceOverlayLayerOrderService({
    visibleOverlayVisible: overlayManager.getVisibleOverlayVisible(),
    invisibleOverlayVisible: overlayManager.getInvisibleOverlayVisible(),
    mainWindow: overlayManager.getMainWindow(),
    invisibleWindow: overlayManager.getInvisibleWindow(),
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
  });
}

async function loadYomitanExtension(): Promise<Extension | null> {
  return loadYomitanExtensionService({
    userDataPath: USER_DATA_PATH,
    getYomitanParserWindow: () => yomitanParserWindow,
    setYomitanParserWindow: (window) => {
      yomitanParserWindow = window;
    },
    setYomitanParserReadyPromise: (promise) => {
      yomitanParserReadyPromise = promise;
    },
    setYomitanParserInitPromise: (promise) => {
      yomitanParserInitPromise = promise;
    },
    setYomitanExtension: (extension) => {
      yomitanExt = extension;
    },
  });
}

function createOverlayWindow(kind: "visible" | "invisible"): BrowserWindow {
  return createOverlayWindowService(
    kind,
    {
      isDev,
      overlayDebugVisualizationEnabled,
      ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
      onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
      setOverlayDebugVisualizationEnabled: (enabled) =>
        setOverlayDebugVisualizationEnabled(enabled),
      isOverlayVisible: (windowKind) =>
        windowKind === "visible"
          ? overlayManager.getVisibleOverlayVisible()
          : overlayManager.getInvisibleOverlayVisible(),
      tryHandleOverlayShortcutLocalFallback: (input) =>
        tryHandleOverlayShortcutLocalFallback(input),
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
  if (overlayRuntimeInitialized) {
    return;
  }
  const result = initializeOverlayRuntimeService(
    {
      backendOverride,
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
        updateVisibleOverlayVisibility();
      },
      updateInvisibleOverlayVisibility: () => {
        updateInvisibleOverlayVisibility();
      },
      getOverlayWindows: () => getOverlayWindows(),
      syncOverlayShortcuts: () => {
        syncOverlayShortcuts();
      },
      setWindowTracker: (tracker) => {
        windowTracker = tracker;
      },
      getResolvedConfig: () => getResolvedConfig(),
      getSubtitleTimingTracker: () => subtitleTimingTracker,
      getMpvClient: () => mpvClient,
      getRuntimeOptionsManager: () => runtimeOptionsManager,
      setAnkiIntegration: (integration) => {
        ankiIntegration = integration as AnkiIntegration | null;
      },
      showDesktopNotification,
      createFieldGroupingCallback: () => createFieldGroupingCallback(),
    },
  );
  overlayManager.setInvisibleOverlayVisible(result.invisibleOverlayVisible);
  overlayRuntimeInitialized = true;
}

function openYomitanSettings(): void {
  openYomitanSettingsWindow(
    {
      yomitanExt,
      getExistingWindow: () => yomitanSettingsWindow,
      setWindow: (window: BrowserWindow | null) => {
        yomitanSettingsWindow = window;
      },
    },
  );
}
function registerGlobalShortcuts(): void {
  registerGlobalShortcutsService(
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

function getOverlayShortcutRuntimeHandlers() {
  return createOverlayShortcutRuntimeHandlers(
    {
    showMpvOsd: (text) => showMpvOsd(text),
    openRuntimeOptions: () => {
      openRuntimeOptionsPalette();
    },
    openJimaku: () => {
      sendToVisibleOverlay("jimaku:open");
    },
    markAudioCard: () => markLastCardAsAudioCard(),
    copySubtitleMultiple: (timeoutMs) => {
      startPendingMultiCopy(timeoutMs);
    },
    copySubtitle: () => {
      copyCurrentSubtitle();
    },
    toggleSecondarySub: () => cycleSecondarySubMode(),
    updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
    triggerFieldGrouping: () => triggerFieldGrouping(),
    triggerSubsync: () => triggerSubsyncFromConfig(),
    mineSentence: () => mineSentenceCard(),
    mineSentenceMultiple: (timeoutMs) => {
      startPendingMineSentenceMultiple(timeoutMs);
    },
    },
  );
}

function tryHandleOverlayShortcutLocalFallback(input: Electron.Input): boolean {
  return runOverlayShortcutLocalFallback(
    input,
    getConfiguredShortcuts(),
    shortcutMatchesInputForLocalFallback,
    getOverlayShortcutRuntimeHandlers().fallbackHandlers,
  );
}

function cycleSecondarySubMode(): void {
  cycleSecondarySubModeService(
    {
      getSecondarySubMode: () => secondarySubMode,
      setSecondarySubMode: (mode: SecondarySubMode) => {
        secondarySubMode = mode;
      },
      getLastSecondarySubToggleAtMs: () => lastSecondarySubToggleAtMs,
      setLastSecondarySubToggleAtMs: (timestampMs: number) => {
        lastSecondarySubToggleAtMs = timestampMs;
      },
      broadcastSecondarySubMode: (mode: SecondarySubMode) => {
        broadcastToOverlayWindows("secondary-subtitle:mode", mode);
      },
      showMpvOsd: (text: string) => showMpvOsd(text),
    },
  );
}

function showMpvOsd(text: string): void {
  showMpvOsdRuntimeService(
    mpvClient,
    text,
    (line) => {
      console.log(line);
    },
  );
}

const numericShortcutRuntime = createNumericShortcutRuntimeService({
  globalShortcut,
  showMpvOsd: (text) => showMpvOsd(text),
  setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
  clearTimer: (timer) => clearTimeout(timer),
});
const multiCopySession = numericShortcutRuntime.createSession();
const mineSentenceSession = numericShortcutRuntime.createSession();

function getSubsyncRuntimeDeps() {
  return {
    getMpvClient: () => mpvClient,
    getResolvedSubsyncConfig: () => getSubsyncConfig(getResolvedConfig().subsync),
    isSubsyncInProgress: () => subsyncInProgress,
    setSubsyncInProgress: (inProgress: boolean) => {
      subsyncInProgress = inProgress;
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
    openManualPicker: (payload: SubsyncManualPayload) => {
      sendToVisibleOverlay("subsync:open-manual", payload, {
        restoreOnModalClose: "subsync",
      });
    },
  };
}

async function triggerSubsyncFromConfig(): Promise<void> {
  await triggerSubsyncFromConfigRuntimeService(getSubsyncRuntimeDeps());
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
  handleMultiCopyDigitService(
    count,
    {
      subtitleTimingTracker,
      writeClipboardText: (text) => clipboard.writeText(text),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

function copyCurrentSubtitle(): void {
  copyCurrentSubtitleService(
    {
      subtitleTimingTracker,
      writeClipboardText: (text) => clipboard.writeText(text),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function updateLastCardFromClipboard(): Promise<void> {
  await updateLastCardFromClipboardService(
    {
      ankiIntegration,
      readClipboardText: () => clipboard.readText(),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function triggerFieldGrouping(): Promise<void> {
  await triggerFieldGroupingService(
    {
      ankiIntegration,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function markLastCardAsAudioCard(): Promise<void> {
  await markLastCardAsAudioCardService(
    {
      ankiIntegration,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function mineSentenceCard(): Promise<void> {
  await mineSentenceCardService(
    {
      ankiIntegration,
      mpvClient,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
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
  handleMineSentenceDigitService(
    count,
    {
      subtitleTimingTracker,
      ankiIntegration,
      getCurrentSecondarySubText: () =>
        mpvClient?.currentSecondarySubText || undefined,
      showMpvOsd: (text) => showMpvOsd(text),
      logError: (message, err) => {
        console.error(message, err);
      },
    },
  );
}

function registerOverlayShortcuts(): void {
  shortcutsRegistered = registerOverlayShortcutsService(
    getConfiguredShortcuts(),
    getOverlayShortcutRuntimeHandlers().overlayHandlers,
  );
}

function getOverlayShortcutLifecycleDeps() {
  return {
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getOverlayHandlers: () => getOverlayShortcutRuntimeHandlers().overlayHandlers,
    cancelPendingMultiCopy: () => cancelPendingMultiCopy(),
    cancelPendingMineSentenceMultiple: () => cancelPendingMineSentenceMultiple(),
  };
}

function unregisterOverlayShortcuts(): void {
  shortcutsRegistered = unregisterOverlayShortcutsRuntimeService(
    shortcutsRegistered,
    getOverlayShortcutLifecycleDeps(),
  );
}

function shouldOverlayShortcutsBeActive(): boolean { return overlayRuntimeInitialized; }
function syncOverlayShortcuts(): void {
  shortcutsRegistered = syncOverlayShortcutsRuntimeService(
    shouldOverlayShortcutsBeActive(),
    shortcutsRegistered,
    getOverlayShortcutLifecycleDeps(),
  );
}
function refreshOverlayShortcuts(): void {
  shortcutsRegistered = refreshOverlayShortcutsRuntimeService(
    shouldOverlayShortcutsBeActive(),
    shortcutsRegistered,
    getOverlayShortcutLifecycleDeps(),
  );
}

function updateVisibleOverlayVisibility(): void {
  updateVisibleOverlayVisibilityService(
    {
      visibleOverlayVisible: overlayManager.getVisibleOverlayVisible(),
      mainWindow: overlayManager.getMainWindow(),
      windowTracker,
      trackerNotReadyWarningShown,
      setTrackerNotReadyWarningShown: (shown) => {
        trackerNotReadyWarningShown = shown;
      },
      shouldBindVisibleOverlayToMpvSubVisibility:
        shouldBindVisibleOverlayToMpvSubVisibility(),
      previousSecondarySubVisibility,
      setPreviousSecondarySubVisibility: (value) => {
        previousSecondarySubVisibility = value;
      },
      mpvConnected: Boolean(mpvClient && mpvClient.connected),
      mpvSend: (payload) => {
        if (!mpvClient) return;
        mpvClient.send(payload);
      },
      secondarySubVisibilityRequestId: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
      updateVisibleOverlayBounds: (geometry) => updateVisibleOverlayBounds(geometry),
      ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
      enforceOverlayLayerOrder: () => enforceOverlayLayerOrder(),
      syncOverlayShortcuts: () => syncOverlayShortcuts(),
    },
  );
}

function updateInvisibleOverlayVisibility(): void {
  updateInvisibleOverlayVisibilityService(
    {
      invisibleWindow: overlayManager.getInvisibleWindow(),
      visibleOverlayVisible: overlayManager.getVisibleOverlayVisible(),
      invisibleOverlayVisible: overlayManager.getInvisibleOverlayVisible(),
      windowTracker,
      updateInvisibleOverlayBounds: (geometry) => updateInvisibleOverlayBounds(geometry),
      ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
      enforceOverlayLayerOrder: () => enforceOverlayLayerOrder(),
      syncOverlayShortcuts: () => syncOverlayShortcuts(),
    },
  );
}

function syncInvisibleOverlayMousePassthrough(): void {
  syncInvisibleOverlayMousePassthroughService({
    hasInvisibleWindow: () => {
      const invisibleWindow = overlayManager.getInvisibleWindow();
      return Boolean(invisibleWindow && !invisibleWindow.isDestroyed());
    },
    setIgnoreMouseEvents: (ignore, extra) => {
      const invisibleWindow = overlayManager.getInvisibleWindow();
      if (!invisibleWindow || invisibleWindow.isDestroyed()) return;
      invisibleWindow.setIgnoreMouseEvents(ignore, extra);
    },
    visibleOverlayVisible: overlayManager.getVisibleOverlayVisible(),
    invisibleOverlayVisible: overlayManager.getInvisibleOverlayVisible(),
  });
}

function setVisibleOverlayVisible(visible: boolean): void {
  setVisibleOverlayVisibleService({
    visible,
    setVisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setVisibleOverlayVisible(nextVisible);
    },
    updateVisibleOverlayVisibility: () => updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () => updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      syncInvisibleOverlayMousePassthrough(),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      shouldBindVisibleOverlayToMpvSubVisibility(),
    isMpvConnected: () => Boolean(mpvClient && mpvClient.connected),
    setMpvSubVisibility: (mpvSubVisible) => {
      setMpvSubVisibilityRuntimeService(mpvClient, mpvSubVisible);
    },
  });
}

function setInvisibleOverlayVisible(visible: boolean): void {
  setInvisibleOverlayVisibleService({
    visible,
    setInvisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setInvisibleOverlayVisible(nextVisible);
    },
    updateInvisibleOverlayVisibility: () => updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      syncInvisibleOverlayMousePassthrough(),
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
  if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
  restoreVisibleOverlayOnModalClose.delete(modal);
  if (restoreVisibleOverlayOnModalClose.size === 0) {
    setVisibleOverlayVisible(false);
  }
}

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcService(
    command,
    {
      specialCommands: SPECIAL_COMMANDS,
      triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
      openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
      runtimeOptionsCycle: (id, direction) => {
        if (!runtimeOptionsManager) {
          return { ok: false, error: "Runtime options manager unavailable" };
        }
        return applyRuntimeOptionResultRuntimeService(
          runtimeOptionsManager.cycleOption(id, direction),
          (text) => showMpvOsd(text),
        );
      },
      showMpvOsd: (text) => showMpvOsd(text),
      mpvReplaySubtitle: () => replayCurrentSubtitleRuntimeService(mpvClient),
      mpvPlayNextSubtitle: () => playNextSubtitleRuntimeService(mpvClient),
      mpvSendCommand: (rawCommand) =>
        sendMpvCommandRuntimeService(mpvClient, rawCommand),
      isMpvConnected: () => Boolean(mpvClient && mpvClient.connected),
      hasRuntimeOptionsManager: () => runtimeOptionsManager !== null,
    },
  );
}

async function runSubsyncManualFromIpc(
  request: SubsyncManualRunRequest,
): Promise<SubsyncResult> {
  return runSubsyncManualFromIpcRuntimeService(request, getSubsyncRuntimeDeps());
}

const runtimeOptionsIpcDeps = {
  setRuntimeOption: (id: string, value: unknown) =>
    setRuntimeOptionFromIpcRuntimeService(
      runtimeOptionsManager,
      id as RuntimeOptionId,
      value as RuntimeOptionValue,
      (text) => showMpvOsd(text),
    ),
  cycleRuntimeOption: (id: string, direction: 1 | -1) =>
    cycleRuntimeOptionFromIpcRuntimeService(
      runtimeOptionsManager,
      id as RuntimeOptionId,
      direction,
      (text) => showMpvOsd(text),
    ),
};

registerIpcHandlersService(
  createIpcDepsRuntimeService({
    getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisibility: () => overlayManager.getVisibleOverlayVisible(),
    getInvisibleOverlayVisibility: () =>
      overlayManager.getInvisibleOverlayVisible(),
    onOverlayModalClosed: (modal) =>
      handleOverlayModalClosed(modal as OverlayHostedModal),
    openYomitanSettings: () => openYomitanSettings(),
    quitApp: () => app.quit(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    tokenizeCurrentSubtitle: () => tokenizeSubtitle(currentSubText),
    getCurrentSubtitleAss: () => currentSubAssText,
    getMpvSubtitleRenderMetrics: () => mpvSubtitleRenderMetrics,
    getSubtitlePosition: () => loadSubtitlePosition(),
    getSubtitleStyle: () => getResolvedConfig().subtitleStyle ?? null,
    saveSubtitlePosition: (position) =>
      saveSubtitlePosition(position as SubtitlePosition),
    getMecabTokenizer: () => mecabTokenizer,
    handleMpvCommand: (command) => handleMpvCommandFromIpc(command),
    getKeybindings: () => keybindings,
    getSecondarySubMode: () => secondarySubMode,
    getMpvClient: () => mpvClient,
    runSubsyncManual: (request) =>
      runSubsyncManualFromIpc(request as SubsyncManualRunRequest),
    getAnkiConnectStatus: () => ankiIntegration !== null,
    getRuntimeOptions: () => getRuntimeOptionsState(),
    setRuntimeOption: runtimeOptionsIpcDeps.setRuntimeOption,
    cycleRuntimeOption: runtimeOptionsIpcDeps.cycleRuntimeOption,
    reportOverlayContentBounds: (payload) => {
      overlayContentMeasurementStore.report(payload);
    },
  }),
);

registerAnkiJimakuIpcRuntimeService(
  {
    patchAnkiConnectEnabled: (enabled) => {
      configService.patchRawConfig({ ankiConnect: { enabled } });
    },
    getResolvedConfig: () => getResolvedConfig(),
    getRuntimeOptionsManager: () => runtimeOptionsManager,
    getSubtitleTimingTracker: () => subtitleTimingTracker,
    getMpvClient: () => mpvClient,
    getAnkiIntegration: () => ankiIntegration,
    setAnkiIntegration: (integration) => {
      ankiIntegration = integration;
    },
    showDesktopNotification,
    createFieldGroupingCallback: () => createFieldGroupingCallback(),
    broadcastRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    getFieldGroupingResolver: () => getFieldGroupingResolver(),
    setFieldGroupingResolver: (resolver) => setFieldGroupingResolver(resolver),
    parseMediaInfo: (mediaPath) => parseMediaInfo(mediaPath),
    getCurrentMediaPath: () => currentMediaPath,
    jimakuFetchJson: (endpoint, query) => jimakuFetchJson(endpoint, query),
    getJimakuMaxEntryResults: () => getJimakuMaxEntryResults(),
    getJimakuLanguagePreference: () => getJimakuLanguagePreference(),
    resolveJimakuApiKey: () => resolveJimakuApiKey(),
    isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath),
    downloadToFile: (url, destPath, headers) =>
      downloadToFile(url, destPath, headers),
  },
);
