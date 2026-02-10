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
  screen,
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
import * as http from "http";
import * as https from "https";
import * as os from "os";
import * as fs from "fs";
import * as crypto from "crypto";
import { MecabTokenizer } from "./mecab-tokenizer";
import { mergeTokens } from "./token-merger";
import { BaseWindowTracker } from "./window-trackers";
import {
  JimakuApiResponse,
  JimakuDownloadResult,
  JimakuMediaInfo,
  JimakuConfig,
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
  CliArgs,
  CliCommandSource,
  hasExplicitCommand,
  parseArgs,
  shouldStartApp,
} from "./cli/args";
import { printHelp } from "./cli/help";
import { generateDefaultConfigFile } from "./core/utils/config-gen";
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
} from "./core/utils/electron-backend";
import { asBoolean, asFiniteNumber, asString } from "./core/utils/coerce";
import { resolveKeybindings } from "./core/utils/keybindings";
import { resolveConfiguredShortcuts } from "./core/utils/shortcut-config";
import { TexthookerService } from "./core/services/texthooker-service";
import {
  hasMpvWebsocketPlugin,
  SubtitleWebSocketService,
} from "./core/services/subtitle-ws-service";
import { registerGlobalShortcutsService } from "./core/services/shortcut-service";
import { registerIpcHandlersService } from "./core/services/ipc-service";
import {
  isGlobalShortcutRegisteredSafe,
  shortcutMatchesInputForLocalFallback,
} from "./core/services/shortcut-fallback-service";
import {
  registerOverlayShortcutsService,
} from "./core/services/overlay-shortcut-service";
import { runOverlayShortcutLocalFallback } from "./core/services/overlay-shortcut-fallback-runner";
import { createOverlayShortcutRuntimeHandlers } from "./core/services/overlay-shortcut-runtime-service";
import { createNumericShortcutSessionService } from "./core/services/numeric-shortcut-session-service";
import { handleCliCommandService } from "./core/services/cli-command-service";
import { cycleSecondarySubModeService } from "./core/services/secondary-subtitle-service";
import {
  refreshOverlayShortcutsRuntimeService,
  syncOverlayShortcutsRuntimeService,
  unregisterOverlayShortcutsRuntimeService,
} from "./core/services/overlay-shortcut-lifecycle-service";
import {
  copyCurrentSubtitleService,
  handleMineSentenceDigitService,
  handleMultiCopyDigitService,
  markLastCardAsAudioCardService,
  mineSentenceCardService,
  triggerFieldGroupingService,
  updateLastCardFromClipboardService,
} from "./core/services/mining-runtime-service";
import { startAppLifecycleService } from "./core/services/app-lifecycle-service";
import {
  playNextSubtitleRuntimeService,
  replayCurrentSubtitleRuntimeService,
  sendMpvCommandRuntimeService,
  setMpvSubVisibilityRuntimeService,
  showMpvOsdRuntimeService,
} from "./core/services/mpv-runtime-service";
import {
  applyRuntimeOptionResultRuntimeService,
  cycleRuntimeOptionFromIpcRuntimeService,
  setRuntimeOptionFromIpcRuntimeService,
} from "./core/services/runtime-options-runtime-service";
import {
  getInitialInvisibleOverlayVisibilityService,
  isAutoUpdateEnabledRuntimeService,
  shouldAutoInitializeOverlayRuntimeFromConfigService,
  shouldBindVisibleOverlayToMpvSubVisibilityService,
} from "./core/services/runtime-config-service";
import { showDesktopNotification } from "./core/utils/notification";
import { openYomitanSettingsWindow } from "./core/services/yomitan-settings-service";
import { tokenizeSubtitleService } from "./core/services/tokenizer-service";
import { loadYomitanExtensionService } from "./core/services/yomitan-extension-loader-service";
import {
  getJimakuLanguagePreferenceService,
  getJimakuMaxEntryResultsService,
  jimakuFetchJsonService,
  resolveJimakuApiKeyService,
} from "./core/services/jimaku-runtime-service";
import {
  loadSubtitlePositionService,
  saveSubtitlePositionService,
  updateCurrentMediaPathService,
} from "./core/services/subtitle-position-service";
import {
  createOverlayWindowService,
  enforceOverlayLayerOrderService,
  ensureOverlayWindowLevelService,
  updateOverlayBoundsService,
} from "./core/services/overlay-window-service";
import { initializeOverlayRuntimeService } from "./core/services/overlay-runtime-init-service";
import {
  syncInvisibleOverlayMousePassthroughService,
} from "./core/services/overlay-visibility-runtime-service";
import {
  setInvisibleOverlayVisibleRuntimeFacadeService,
  setVisibleOverlayVisibleRuntimeFacadeService,
  toggleInvisibleOverlayRuntimeFacadeService,
  toggleVisibleOverlayRuntimeFacadeService,
} from "./core/services/overlay-visibility-facade-service";
import {
  MpvIpcClient,
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
} from "./core/services/mpv-service";
import { applyMpvSubtitleRenderMetricsPatchService } from "./core/services/mpv-render-metrics-service";
import {
  handleMpvCommandFromIpcService,
} from "./core/services/ipc-command-service";
import {
  handleOverlayModalClosedService,
} from "./core/services/overlay-modal-restore-service";
import {
  createFieldGroupingCallbackRuntimeService,
  sendToVisibleOverlayRuntimeService,
} from "./core/services/overlay-bridge-runtime-service";
import {
  broadcastRuntimeOptionsChangedRuntimeService,
  broadcastToOverlayWindowsRuntimeService,
  getOverlayWindowsRuntimeService,
  setOverlayDebugVisualizationEnabledRuntimeService,
} from "./core/services/overlay-broadcast-runtime-service";
import { runAppReadyRuntimeService } from "./core/services/app-ready-runtime-service";
import { runAppShutdownRuntimeService } from "./core/services/app-shutdown-runtime-service";
import { createMpvIpcClientDepsRuntimeService } from "./core/services/mpv-client-deps-runtime-service";
import { createAppLifecycleDepsRuntimeService } from "./core/services/app-lifecycle-deps-runtime-service";
import { createCliCommandDepsRuntimeService } from "./core/services/cli-command-deps-runtime-service";
import { createIpcDepsRuntimeService } from "./core/services/ipc-deps-runtime-service";
import { createRuntimeOptionsManagerRuntimeService } from "./core/services/runtime-options-manager-runtime-service";
import { createAppLoggingRuntimeService } from "./core/services/app-logging-runtime-service";
import {
  createMecabTokenizerAndCheckRuntimeService,
  createSubtitleTimingTrackerRuntimeService,
} from "./core/services/startup-resource-runtime-service";
import { runGenerateConfigFlowRuntimeService } from "./core/services/config-generation-runtime-service";
import {
  runSubsyncManualFromIpcRuntimeService,
  triggerSubsyncFromConfigRuntimeService,
} from "./core/services/subsync-runtime-service";
import {
  updateInvisibleOverlayVisibilityService,
  updateVisibleOverlayVisibilityService,
} from "./core/services/overlay-visibility-service";
import { registerAnkiJimakuIpcRuntimeService } from "./core/services/anki-jimaku-runtime-service";
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
const CONFIG_DIR = path.join(os.homedir(), ".config", "SubMiner");
const USER_DATA_PATH = CONFIG_DIR;
const configService = new ConfigService(CONFIG_DIR);
const isDev =
  process.argv.includes("--dev") || process.argv.includes("--debug");
const texthookerService = new TexthookerService();
const subtitleWsService = new SubtitleWebSocketService();
const appLogger = createAppLoggingRuntimeService();

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

let mainWindow: BrowserWindow | null = null;
let invisibleWindow: BrowserWindow | null = null;
let yomitanExt: Extension | null = null;
let yomitanSettingsWindow: BrowserWindow | null = null;
let yomitanParserWindow: BrowserWindow | null = null;
let yomitanParserReadyPromise: Promise<void> | null = null;
let yomitanParserInitPromise: Promise<boolean> | null = null;
let mpvClient: MpvIpcClient | null = null;
let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
let currentSubText = "";
let currentSubAssText = "";
let visibleOverlayVisible = false;
let invisibleOverlayVisible = false;
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
const DEFAULT_MPV_SUBTITLE_RENDER_METRICS: MpvSubtitleRenderMetrics = {
  subPos: 100,
  subFontSize: 38,
  subScale: 1,
  subMarginY: 34,
  subMarginX: 19,
  subFont: "sans-serif",
  subSpacing: 0,
  subBold: false,
  subItalic: false,
  subBorderSize: 2.5,
  subShadowOffset: 0,
  subAssOverride: "yes",
  subScaleByWindow: true,
  subUseMargins: true,
  osdHeight: 720,
  osdDimensions: null,
};
let mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics = {
  ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
};

let shortcutsRegistered = false;
let overlayRuntimeInitialized = false;
let fieldGroupingResolver: ((choice: KikuFieldGroupingChoice) => void) | null =
  null;
let runtimeOptionsManager: RuntimeOptionsManager | null = null;
let trackerNotReadyWarningShown = false;
let overlayDebugVisualizationEnabled = false;
type OverlayHostedModal = "runtime-options" | "subsync";
const restoreVisibleOverlayOnModalClose = new Set<OverlayHostedModal>();

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, "subtitle-positions");

function getRuntimeOptionsState(): RuntimeOptionState[] { if (!runtimeOptionsManager) return []; return runtimeOptionsManager.listOptions(); }

function getOverlayWindows(): BrowserWindow[] {
  return getOverlayWindowsRuntimeService({ mainWindow, invisibleWindow });
}

function broadcastToOverlayWindows(channel: string, ...args: unknown[]): void {
  broadcastToOverlayWindowsRuntimeService(getOverlayWindows(), channel, ...args);
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

const initialArgs = parseArgs(process.argv);
if (initialArgs.logLevel) {
  process.env.SUBMINER_LOG_LEVEL = initialArgs.logLevel;
} else if (initialArgs.verbose) {
  process.env.SUBMINER_LOG_LEVEL = "debug";
}

forceX11Backend(initialArgs);
enforceUnsupportedWaylandMode(initialArgs);

let mpvSocketPath = initialArgs.socketPath ?? getDefaultSocketPath();
let texthookerPort = initialArgs.texthookerPort ?? DEFAULT_TEXTHOOKER_PORT;
const backendOverride = initialArgs.backend ?? null;
const autoStartOverlay = initialArgs.autoStartOverlay;
const texthookerOnlyMode = initialArgs.texthooker;

if (
  !runGenerateConfigFlowRuntimeService(initialArgs, {
    shouldStartApp: (args) => shouldStartApp(args),
    generateConfig: async (args) =>
      generateDefaultConfigFile(args, {
        configDir: CONFIG_DIR,
        defaultConfig: DEFAULT_CONFIG,
        generateTemplate: (config) => generateConfigTemplate(config as never),
      }),
    onSuccess: (exitCode) => {
      process.exitCode = exitCode;
      app.quit();
    },
    onError: (error) => {
      console.error(`Failed to generate config: ${error.message}`);
      process.exitCode = 1;
      app.quit();
    },
  })
) {
  startAppLifecycleService(initialArgs, createAppLifecycleDepsRuntimeService({
    app,
    platform: process.platform,
    shouldStartApp: (args) => shouldStartApp(args),
    parseArgs: (argv) => parseArgs(argv),
    handleCliCommand: (args, source) => handleCliCommand(args, source),
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
            createMpvIpcClientDepsRuntimeService({
              getResolvedConfig: () => getResolvedConfig(),
              autoStartOverlay,
              setOverlayVisible: (visible) => setOverlayVisible(visible),
              shouldBindVisibleOverlayToMpvSubVisibility: () =>
                shouldBindVisibleOverlayToMpvSubVisibility(),
              isVisibleOverlayVisible: () => visibleOverlayVisible,
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
              broadcastToOverlayWindows: (channel, ...args) => {
                broadcastToOverlayWindows(channel, ...args);
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
            }),
          );
        },
        reloadConfig: () => {
          configService.reloadConfig();
        },
        getResolvedConfig: () => getResolvedConfig(),
        getConfigWarnings: () => configService.getWarnings(),
        logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
        initRuntimeOptionsManager: () => {
          runtimeOptionsManager = createRuntimeOptionsManagerRuntimeService({
            getAnkiConfig: () => configService.getConfig().ankiConnect,
            applyAnkiPatch: (patch) => {
              if (ankiIntegration) {
                ankiIntegration.applyRuntimeConfigPatch(patch);
              }
            },
            onOptionsChanged: () => {
              broadcastRuntimeOptionsChanged();
              refreshOverlayShortcuts();
            },
          });
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
        createMecabTokenizerAndCheck: async () =>
          createMecabTokenizerAndCheckRuntimeService({
            createMecabTokenizer: () => new MecabTokenizer(),
            setMecabTokenizer: (tokenizer) => {
              mecabTokenizer = tokenizer;
            },
          }),
        createSubtitleTimingTracker: () =>
          createSubtitleTimingTrackerRuntimeService({
            createSubtitleTimingTracker: () => new SubtitleTimingTracker(),
            setSubtitleTimingTracker: (tracker) => {
              subtitleTimingTracker = tracker;
            },
          }),
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
      runAppShutdownRuntimeService({
        unregisterAllGlobalShortcuts: () => {
          globalShortcut.unregisterAll();
        },
        stopSubtitleWebsocket: () => {
          subtitleWsService.stop();
        },
        stopTexthookerService: () => {
          texthookerService.stop();
        },
        destroyYomitanParserWindow: () => {
          if (yomitanParserWindow && !yomitanParserWindow.isDestroyed()) {
            yomitanParserWindow.destroy();
          }
          yomitanParserWindow = null;
        },
        clearYomitanParserPromises: () => {
          yomitanParserReadyPromise = null;
          yomitanParserInitPromise = null;
        },
        stopWindowTracker: () => {
          if (windowTracker) {
            windowTracker.stop();
          }
        },
        destroyMpvSocket: () => {
          if (mpvClient && mpvClient.socket) {
            mpvClient.socket.destroy();
          }
        },
        clearReconnectTimer: () => {
          if (reconnectTimer) {
            clearTimeout(reconnectTimer);
          }
        },
        destroySubtitleTimingTracker: () => {
          if (subtitleTimingTracker) {
            subtitleTimingTracker.destroy();
          }
        },
        destroyAnkiIntegration: () => {
          if (ankiIntegration) {
            ankiIntegration.destroy();
          }
        },
      });
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
}

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
        shell.openExternal(url);
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
      hasMainWindow: () => Boolean(mainWindow),
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
  return tokenizeSubtitleService(text, {
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
    tokenizeWithMecab: async (tokenizeText) => {
      if (!mecabTokenizer) {
        return null;
      }
      const rawTokens = await mecabTokenizer.tokenize(tokenizeText);
      if (!rawTokens || rawTokens.length === 0) {
        return null;
      }
      return mergeTokens(rawTokens);
    },
  });
}

function updateOverlayBounds(geometry: WindowGeometry): void {
  updateOverlayBoundsService(geometry, () => getOverlayWindows());
}

function ensureOverlayWindowLevel(window: BrowserWindow): void {
  ensureOverlayWindowLevelService(window);
}

function enforceOverlayLayerOrder(): void {
  enforceOverlayLayerOrderService({
    visibleOverlayVisible,
    invisibleOverlayVisible,
    mainWindow,
    invisibleWindow,
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
  return createOverlayWindowService(kind, {
    isDev,
    overlayDebugVisualizationEnabled,
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
    onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    setOverlayDebugVisualizationEnabled: (enabled) =>
      setOverlayDebugVisualizationEnabled(enabled),
    isOverlayVisible: (windowKind) =>
      windowKind === "visible" ? visibleOverlayVisible : invisibleOverlayVisible,
    tryHandleOverlayShortcutLocalFallback: (input) =>
      tryHandleOverlayShortcutLocalFallback(input),
    onWindowClosed: (windowKind) => {
      if (windowKind === "visible") {
        mainWindow = null;
      } else {
        invisibleWindow = null;
      }
    },
  });
}

function createMainWindow(): BrowserWindow { mainWindow = createOverlayWindow("visible"); return mainWindow; }
function createInvisibleWindow(): BrowserWindow { invisibleWindow = createOverlayWindow("invisible"); return invisibleWindow; }

function initializeOverlayRuntime(): void {
  if (overlayRuntimeInitialized) {
    return;
  }
  const result = initializeOverlayRuntimeService({
    backendOverride,
    getInitialInvisibleOverlayVisibility: () => getInitialInvisibleOverlayVisibility(),
    createMainWindow: () => {
      createMainWindow();
    },
    createInvisibleWindow: () => {
      createInvisibleWindow();
    },
    registerGlobalShortcuts: () => {
      registerGlobalShortcuts();
    },
    updateOverlayBounds: (geometry) => {
      updateOverlayBounds(geometry);
    },
    isVisibleOverlayVisible: () => visibleOverlayVisible,
    isInvisibleOverlayVisible: () => invisibleOverlayVisible,
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
  });
  invisibleOverlayVisible = result.invisibleOverlayVisible;
  overlayRuntimeInitialized = true;
}

function openYomitanSettings(): void { openYomitanSettingsWindow({ yomitanExt, getExistingWindow: () => yomitanSettingsWindow, setWindow: (window) => (yomitanSettingsWindow = window) }); }
function registerGlobalShortcuts(): void { registerGlobalShortcutsService({ shortcuts: getConfiguredShortcuts(), onToggleVisibleOverlay: () => toggleVisibleOverlay(), onToggleInvisibleOverlay: () => toggleInvisibleOverlay(), onOpenYomitanSettings: () => openYomitanSettings(), isDev, getMainWindow: () => mainWindow }); }

function getConfiguredShortcuts() { return resolveConfiguredShortcuts(getResolvedConfig(), DEFAULT_CONFIG); }

function getOverlayShortcutRuntimeHandlers() {
  return createOverlayShortcutRuntimeHandlers({
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
  });
}

function tryHandleOverlayShortcutLocalFallback(input: Electron.Input): boolean {
  const shortcuts = getConfiguredShortcuts();
  const handlers = getOverlayShortcutRuntimeHandlers();
  return runOverlayShortcutLocalFallback(
    input,
    shortcuts,
    shortcutMatchesInputForLocalFallback,
    handlers.fallbackHandlers,
  );
}

function cycleSecondarySubMode(): void {
  cycleSecondarySubModeService({
    getSecondarySubMode: () => secondarySubMode,
    setSecondarySubMode: (mode) => {
      secondarySubMode = mode;
    },
    getLastSecondarySubToggleAtMs: () => lastSecondarySubToggleAtMs,
    setLastSecondarySubToggleAtMs: (timestampMs) => {
      lastSecondarySubToggleAtMs = timestampMs;
    },
    broadcastSecondarySubMode: (mode) => {
      broadcastToOverlayWindows("secondary-subtitle:mode", mode);
    },
    showMpvOsd: (text) => showMpvOsd(text),
  });
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

const multiCopySession = createNumericShortcutSessionService({
  registerShortcut: (accelerator, handler) =>
    globalShortcut.register(accelerator, handler),
  unregisterShortcut: (accelerator) => globalShortcut.unregister(accelerator),
  setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
  clearTimer: (timer) => clearTimeout(timer),
  showMpvOsd: (text) => showMpvOsd(text),
});

const mineSentenceSession = createNumericShortcutSessionService({
  registerShortcut: (accelerator, handler) =>
    globalShortcut.register(accelerator, handler),
  unregisterShortcut: (accelerator) => globalShortcut.unregister(accelerator),
  setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
  clearTimer: (timer) => clearTimeout(timer),
  showMpvOsd: (text) => showMpvOsd(text),
});

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
  handleMultiCopyDigitService(count, {
    subtitleTimingTracker,
    writeClipboardText: (text) => clipboard.writeText(text),
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

function copyCurrentSubtitle(): void {
  copyCurrentSubtitleService({
    subtitleTimingTracker,
    writeClipboardText: (text) => clipboard.writeText(text),
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function updateLastCardFromClipboard(): Promise<void> {
  await updateLastCardFromClipboardService({
    ankiIntegration,
    readClipboardText: () => clipboard.readText(),
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function triggerFieldGrouping(): Promise<void> {
  await triggerFieldGroupingService({
    ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function markLastCardAsAudioCard(): Promise<void> {
  await markLastCardAsAudioCardService({
    ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

async function mineSentenceCard(): Promise<void> {
  await mineSentenceCardService({
    ankiIntegration,
    mpvClient,
    showMpvOsd: (text) => showMpvOsd(text),
  });
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
  handleMineSentenceDigitService(count, {
    subtitleTimingTracker,
    ankiIntegration,
    getCurrentSecondarySubText: () => mpvClient?.currentSecondarySubText || undefined,
    showMpvOsd: (text) => showMpvOsd(text),
    logError: (message, err) => {
      console.error(message, err);
    },
  });
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
  updateVisibleOverlayVisibilityService({
    visibleOverlayVisible,
    mainWindow,
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
    updateOverlayBounds: (geometry) => updateOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
    enforceOverlayLayerOrder: () => enforceOverlayLayerOrder(),
    syncOverlayShortcuts: () => syncOverlayShortcuts(),
  });
}

function updateInvisibleOverlayVisibility(): void {
  updateInvisibleOverlayVisibilityService({
    invisibleWindow,
    visibleOverlayVisible,
    invisibleOverlayVisible,
    windowTracker,
    updateOverlayBounds: (geometry) => updateOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
    enforceOverlayLayerOrder: () => enforceOverlayLayerOrder(),
    syncOverlayShortcuts: () => syncOverlayShortcuts(),
  });
}

function syncInvisibleOverlayMousePassthrough(): void {
  syncInvisibleOverlayMousePassthroughService({
    hasInvisibleWindow: () => Boolean(invisibleWindow && !invisibleWindow.isDestroyed()),
    setIgnoreMouseEvents: (ignore, extra) => {
      if (!invisibleWindow || invisibleWindow.isDestroyed()) return;
      invisibleWindow.setIgnoreMouseEvents(ignore, extra);
    },
    visibleOverlayVisible,
    invisibleOverlayVisible,
  });
}

function setVisibleOverlayVisible(visible: boolean): void {
  setVisibleOverlayVisibleRuntimeFacadeService(visible, getOverlayVisibilityFacadeDeps());
}

function setInvisibleOverlayVisible(visible: boolean): void {
  setInvisibleOverlayVisibleRuntimeFacadeService(visible, getOverlayVisibilityFacadeDeps());
}

function getOverlayVisibilityFacadeDeps() {
  return {
    getVisibleOverlayVisible: () => visibleOverlayVisible,
    getInvisibleOverlayVisible: () => invisibleOverlayVisible,
    setVisibleOverlayVisibleState: (nextVisible: boolean) => {
      visibleOverlayVisible = nextVisible;
    },
    setInvisibleOverlayVisibleState: (nextVisible: boolean) => {
      invisibleOverlayVisible = nextVisible;
    },
    updateVisibleOverlayVisibility: () => updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () => updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      syncInvisibleOverlayMousePassthrough(),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      shouldBindVisibleOverlayToMpvSubVisibility(),
    isMpvConnected: () => Boolean(mpvClient && mpvClient.connected),
    setMpvSubVisibility: (mpvSubVisible: boolean) => {
      setMpvSubVisibilityRuntimeService(mpvClient, mpvSubVisible);
    },
  };
}

function toggleVisibleOverlay(): void {
  toggleVisibleOverlayRuntimeFacadeService(getOverlayVisibilityFacadeDeps());
}
function toggleInvisibleOverlay(): void {
  toggleInvisibleOverlayRuntimeFacadeService(getOverlayVisibilityFacadeDeps());
}
function setOverlayVisible(visible: boolean): void { setVisibleOverlayVisible(visible); }
function toggleOverlay(): void { toggleVisibleOverlay(); }
function handleOverlayModalClosed(modal: OverlayHostedModal): void {
  handleOverlayModalClosedService(
    restoreVisibleOverlayOnModalClose,
    modal,
    (visible) => setVisibleOverlayVisible(visible),
  );
}

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcService(command, {
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
  });
}

async function runSubsyncManualFromIpc(
  request: SubsyncManualRunRequest,
): Promise<SubsyncResult> {
  return runSubsyncManualFromIpcRuntimeService(request, getSubsyncRuntimeDeps());
}

registerIpcHandlersService(
  createIpcDepsRuntimeService({
    getInvisibleWindow: () => invisibleWindow,
    getMainWindow: () => mainWindow,
    getVisibleOverlayVisibility: () => visibleOverlayVisible,
    getInvisibleOverlayVisibility: () => invisibleOverlayVisible,
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
    setRuntimeOption: (id, value) =>
      setRuntimeOptionFromIpcRuntimeService(
        runtimeOptionsManager,
        id as RuntimeOptionId,
        value as RuntimeOptionValue,
        (text) => showMpvOsd(text),
      ),
    cycleRuntimeOption: (id, direction) =>
      cycleRuntimeOptionFromIpcRuntimeService(
        runtimeOptionsManager,
        id as RuntimeOptionId,
        direction,
        (text) => showMpvOsd(text),
      ),
  }),
);

/**
 * Create and show a desktop notification with robust icon handling.
 * Supports both file paths (preferred on Linux/Wayland) and data URLs (fallback).
 */
function createFieldGroupingCallback() {
  return createFieldGroupingCallbackRuntimeService({
    getVisibleOverlayVisible: () => visibleOverlayVisible,
    getInvisibleOverlayVisible: () => invisibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    setInvisibleOverlayVisible: (visible) => setInvisibleOverlayVisible(visible),
    getResolver: () => fieldGroupingResolver,
    setResolver: (resolver) => {
      fieldGroupingResolver = resolver;
    },
    sendToVisibleOverlay: (
      channel,
      payload,
      runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
    ) =>
      sendToVisibleOverlay(channel, payload, runtimeOptions),
  });
}

function sendToVisibleOverlay(channel: string, payload?: unknown, options?: { restoreOnModalClose?: OverlayHostedModal }): boolean {
  return sendToVisibleOverlayRuntimeService({
    mainWindow,
    visibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    channel,
    payload,
    restoreOnModalClose: options?.restoreOnModalClose,
    restoreVisibleOverlayOnModalClose,
  });
}

registerAnkiJimakuIpcRuntimeService({
  patchAnkiConnectEnabled: (enabled) => { configService.patchRawConfig({ ankiConnect: { enabled } }); },
  getResolvedConfig: () => getResolvedConfig(), getRuntimeOptionsManager: () => runtimeOptionsManager, getSubtitleTimingTracker: () => subtitleTimingTracker, getMpvClient: () => mpvClient, getAnkiIntegration: () => ankiIntegration, setAnkiIntegration: (integration) => { ankiIntegration = integration; },
  showDesktopNotification, createFieldGroupingCallback: () => createFieldGroupingCallback(), broadcastRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
  getFieldGroupingResolver: () => fieldGroupingResolver, setFieldGroupingResolver: (resolver) => { fieldGroupingResolver = resolver; },
  parseMediaInfo: (mediaPath) => parseMediaInfo(mediaPath), getCurrentMediaPath: () => currentMediaPath, jimakuFetchJson: (endpoint, query) => jimakuFetchJson(endpoint, query),
  getJimakuMaxEntryResults: () => getJimakuMaxEntryResults(), getJimakuLanguagePreference: () => getJimakuLanguagePreference(), resolveJimakuApiKey: () => resolveJimakuApiKey(),
  isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath), downloadToFile: (url, destPath, headers) => downloadToFile(url, destPath, headers),
});
