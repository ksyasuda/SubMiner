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
import type {
  AppLifecycleDepsRuntimeOptions,
} from "./core/services/app-lifecycle-service";
import type { CliCommandDepsRuntimeOptions } from "./core/services/cli-command-service";
import type { HandleMpvCommandFromIpcOptions } from "./core/services/ipc-command-service";
import type { IpcDepsRuntimeOptions } from "./core/services/ipc-service";
import type { AnkiJimakuIpcRuntimeOptions } from "./core/services/anki-jimaku-service";
import type { StartupBootstrapRuntimeDeps } from "./core/services/startup-service";

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

const overlayManager = createOverlayManagerService();
const overlayContentMeasurementStore = createOverlayContentMeasurementStoreService({
  now: () => Date.now(),
  warn: (message: string) => {
    console.warn(message);
  },
});
type OverlayHostedModal = "runtime-options" | "subsync" | "jimaku";
type OverlayHostLayer = "visible" | "invisible";
const restoreVisibleOverlayOnModalClose = new Set<OverlayHostedModal>();
const overlayModalAutoShownLayer = new Map<OverlayHostedModal, OverlayHostLayer>();

interface AppState {
  yomitanExt: Extension | null;
  yomitanSettingsWindow: BrowserWindow | null;
  yomitanParserWindow: BrowserWindow | null;
  yomitanParserReadyPromise: Promise<void> | null;
  yomitanParserInitPromise: Promise<boolean> | null;
  mpvClient: MpvIpcClient | null;
  reconnectTimer: ReturnType<typeof setTimeout> | null;
  currentSubText: string;
  currentSubAssText: string;
  windowTracker: BaseWindowTracker | null;
  subtitlePosition: SubtitlePosition | null;
  currentMediaPath: string | null;
  currentMediaTitle: string | null;
  pendingSubtitlePosition: SubtitlePosition | null;
  mecabTokenizer: MecabTokenizer | null;
  keybindings: Keybinding[];
  subtitleTimingTracker: SubtitleTimingTracker | null;
  ankiIntegration: AnkiIntegration | null;
  secondarySubMode: SecondarySubMode;
  lastSecondarySubToggleAtMs: number;
  previousSecondarySubVisibility: boolean | null;
  mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics;
  shortcutsRegistered: boolean;
  overlayRuntimeInitialized: boolean;
  fieldGroupingResolver: ((choice: KikuFieldGroupingChoice) => void) | null;
  fieldGroupingResolverSequence: number;
  runtimeOptionsManager: RuntimeOptionsManager | null;
  trackerNotReadyWarningShown: boolean;
  overlayDebugVisualizationEnabled: boolean;
  subsyncInProgress: boolean;
  initialArgs: CliArgs | null;
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride: string | null;
  autoStartOverlay: boolean;
  texthookerOnlyMode: boolean;
}

const appState: AppState = {
  yomitanExt: null,
  yomitanSettingsWindow: null,
  yomitanParserWindow: null,
  yomitanParserReadyPromise: null,
  yomitanParserInitPromise: null,
  mpvClient: null,
  reconnectTimer: null,
  currentSubText: "",
  currentSubAssText: "",
  windowTracker: null,
  subtitlePosition: null,
  currentMediaPath: null,
  currentMediaTitle: null,
  pendingSubtitlePosition: null,
  mecabTokenizer: null,
  keybindings: [],
  subtitleTimingTracker: null,
  ankiIntegration: null,
  secondarySubMode: "hover",
  lastSecondarySubToggleAtMs: 0,
  previousSecondarySubVisibility: null,
  mpvSubtitleRenderMetrics: {
    ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
  },
  shortcutsRegistered: false,
  overlayRuntimeInitialized: false,
  fieldGroupingResolver: null,
  fieldGroupingResolverSequence: 0,
  runtimeOptionsManager: null,
  trackerNotReadyWarningShown: false,
  overlayDebugVisualizationEnabled: false,
  subsyncInProgress: false,
  initialArgs: null,
  mpvSocketPath: getDefaultSocketPath(),
  texthookerPort: DEFAULT_TEXTHOOKER_PORT,
  backendOverride: null,
  autoStartOverlay: false,
  texthookerOnlyMode: false,
};

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

const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntimeService<OverlayHostedModal>({
  getMainWindow: () => overlayManager.getMainWindow(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
  setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  setInvisibleOverlayVisible: (visible) => setInvisibleOverlayVisible(visible),
  getResolver: () => getFieldGroupingResolver(),
  setResolver: (resolver) => setFieldGroupingResolver(resolver),
  getRestoreVisibleOverlayOnModalClose: () => restoreVisibleOverlayOnModalClose,
  sendToVisibleOverlay: (channel, payload) => {
    sendToActiveOverlayWindow(channel, payload);
    return true;
  },
});
const createFieldGroupingCallback =
  fieldGroupingOverlayRuntime.createFieldGroupingCallback;

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, "subtitle-positions");

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
  broadcastRuntimeOptionsChangedRuntimeService(
    () => getRuntimeOptionsState(),
    (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  );
}

function getTargetOverlayWindow(): {
  window: BrowserWindow;
  layer: OverlayHostLayer;
} | null {
  const visibleMainWindow = overlayManager.getMainWindow();
  const invisibleWindow = overlayManager.getInvisibleWindow();

  if (visibleMainWindow && !visibleMainWindow.isDestroyed()) {
    return { window: visibleMainWindow, layer: "visible" };
  }

  if (invisibleWindow && !invisibleWindow.isDestroyed()) {
    return { window: invisibleWindow, layer: "invisible" };
  }

  return null;
}

function showOverlayWindowForModal(window: BrowserWindow, layer: OverlayHostLayer): void {
  if (layer === "invisible" && typeof window.showInactive === "function") {
    window.showInactive();
  } else {
    window.show();
  }
  if (!window.isFocused()) {
    window.focus();
  }
}

function sendToActiveOverlayWindow(
  channel: string,
  payload?: unknown,
  runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
): void {
  const target = getTargetOverlayWindow();
  if (!target) return;

  const { window: targetWindow, layer } = target;
  const wasVisible = targetWindow.isVisible();
  const restoreOnModalClose = runtimeOptions?.restoreOnModalClose;

  const sendNow = (): void => {
    if (payload === undefined) {
      targetWindow.webContents.send(channel);
    } else {
      targetWindow.webContents.send(channel, payload);
    }
  };

  if (!wasVisible) {
    showOverlayWindowForModal(targetWindow, layer);
  }
  if (!wasVisible && restoreOnModalClose) {
    restoreVisibleOverlayOnModalClose.add(restoreOnModalClose);
    overlayModalAutoShownLayer.set(restoreOnModalClose, layer);
  }

  if (targetWindow.webContents.isLoading()) {
    targetWindow.webContents.once("did-finish-load", () => {
      if (
        targetWindow &&
        !targetWindow.isDestroyed() &&
        !targetWindow.webContents.isLoading()
      ) {
        sendNow();
      }
    });
    return;
  }

  sendNow();
}

function setOverlayDebugVisualizationEnabled(enabled: boolean): void {
  setOverlayDebugVisualizationEnabledRuntimeService(
    appState.overlayDebugVisualizationEnabled,
    enabled,
    (next) => {
      appState.overlayDebugVisualizationEnabled = next;
    },
    (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  );
}

function openRuntimeOptionsPalette(): void {
  sendToActiveOverlayWindow("runtime-options:open", undefined, {
    restoreOnModalClose: "runtime-options",
  });
}

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
    appState.runtimeOptionsManager,
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
  appState.subtitlePosition = loadSubtitlePositionService({
    currentMediaPath: appState.currentMediaPath,
    fallbackPosition: getResolvedConfig().subtitlePosition,
    subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
  });
  return appState.subtitlePosition;
}

function saveSubtitlePosition(position: SubtitlePosition): void {
  appState.subtitlePosition = position;
  saveSubtitlePositionService({
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

function updateCurrentMediaPath(mediaPath: unknown): void {
  if (typeof mediaPath !== "string" || !isRemoteMediaPath(mediaPath)) {
    appState.currentMediaTitle = null;
  }
  updateCurrentMediaPathService({
    mediaPath,
    currentMediaPath: appState.currentMediaPath,
    pendingSubtitlePosition: appState.pendingSubtitlePosition,
    subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
    loadSubtitlePosition: () => loadSubtitlePosition(),
    setCurrentMediaPath: (nextPath) => {
      appState.currentMediaPath = nextPath;
    },
    clearPendingSubtitlePosition: () => {
      appState.pendingSubtitlePosition = null;
    },
    setSubtitlePosition: (position) => {
      appState.subtitlePosition = position;
    },
    broadcastSubtitlePosition: (position) => {
      broadcastToOverlayWindows("subtitle-position:set", position);
    },
  });
}

function updateCurrentMediaTitle(mediaTitle: unknown): void {
  if (typeof mediaTitle === "string") {
    const sanitized = mediaTitle.trim();
    appState.currentMediaTitle = sanitized.length > 0 ? sanitized : null;
    return;
  }
  appState.currentMediaTitle = null;
}

function resolveMediaPathForJimaku(mediaPath: string | null): string | null {
  return mediaPath && isRemoteMediaPath(mediaPath) && appState.currentMediaTitle
    ? appState.currentMediaTitle
    : mediaPath;
}

const startupState = runStartupBootstrapRuntimeService(createStartupBootstrapRuntimeDeps());

appState.initialArgs = startupState.initialArgs;
appState.mpvSocketPath = startupState.mpvSocketPath;
appState.texthookerPort = startupState.texthookerPort;
appState.backendOverride = startupState.backendOverride;
appState.autoStartOverlay = startupState.autoStartOverlay;
appState.texthookerOnlyMode = startupState.texthookerOnlyMode;

function createAppLifecycleRuntimeDeps(): AppLifecycleDepsRuntimeOptions {
  return {
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
          appState.keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);
        },
        createMpvClient: () => {
          appState.mpvClient = createMpvClientRuntimeService();
        },
        reloadConfig: () => {
          configService.reloadConfig();
          appLogger.logInfo(`Using config file: ${configService.getConfigPath()}`);
        },
        getResolvedConfig: () => getResolvedConfig(),
        getConfigWarnings: () => configService.getWarnings(),
        logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
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
        setSecondarySubMode: (mode) => {
          appState.secondarySubMode = mode;
        },
        defaultSecondarySubMode: "hover",
        defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
        hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
        startSubtitleWebsocket: (port) => {
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
        loadYomitanExtension: async () => {
          await loadYomitanExtension();
        },
        texthookerOnlyMode: appState.texthookerOnlyMode,
        shouldAutoInitializeOverlayRuntimeFromConfig: () =>
          shouldAutoInitializeOverlayRuntimeFromConfig(),
        initializeOverlayRuntime: () => initializeOverlayRuntime(),
        handleInitialArgs: () => handleInitialArgs(),
      });
    },
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
      if (appState.ankiIntegration) {
        appState.ankiIntegration.destroy();
      }
    },
    shouldRestoreWindowsOnActivate: () =>
      appState.overlayRuntimeInitialized && BrowserWindow.getAllWindows().length === 0,
    restoreWindowsOnActivate: () => {
      createMainWindow();
      createInvisibleWindow();
      updateVisibleOverlayVisibility();
      updateInvisibleOverlayVisibility();
    },
  };
}

function createStartupBootstrapRuntimeDeps(): StartupBootstrapRuntimeDeps {
  return {
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
      startAppLifecycleService(
        args,
        createAppLifecycleDepsRuntimeService(createAppLifecycleRuntimeDeps()),
      );
    },
  };
}

function handleCliCommand(
  args: CliArgs,
  source: CliCommandSource = "initial",
): void {
  const deps = createCliCommandDepsRuntimeService(createCliCommandRuntimeServiceDeps());
  handleCliCommandService(args, source, deps);
}

function createCliCommandRuntimeServiceDeps(): CliCommandDepsRuntimeOptions {
  return {
    mpv: {
      getSocketPath: () => appState.mpvSocketPath,
      setSocketPath: (socketPath) => {
        appState.mpvSocketPath = socketPath;
      },
      getClient: () => appState.mpvClient,
      showOsd: (text) => showMpvOsd(text),
    },
    texthooker: {
      service: texthookerService,
      getPort: () => appState.texthookerPort,
      setPort: (port) => {
        appState.texthookerPort = port;
      },
      shouldOpenBrowser: () => getResolvedConfig().texthooker?.openBrowser !== false,
      openInBrowser: (url) => {
        void shell.openExternal(url).catch((error) => {
          console.error(`Failed to open browser for texthooker URL: ${url}`, error);
        });
      },
    },
    overlay: {
      isInitialized: () => appState.overlayRuntimeInitialized,
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
  };
}

function handleInitialArgs(): void {
  if (!appState.initialArgs) return;
  handleCliCommand(appState.initialArgs, "initial");
}

function bindMpvClientEventHandlers(mpvClient: MpvIpcClient): void {
  mpvClient.on("media-path-change", ({ path }) => {
    updateCurrentMediaPath(path);
  });
  mpvClient.on("media-title-change", ({ title }) => {
    updateCurrentMediaTitle(title);
  });
  mpvClient.on("subtitle-ass-change", ({ text }) => {
    appState.currentSubAssText = text;
  });
  mpvClient.on("subtitle-change", ({ text }) => {
    appState.currentSubText = text;
  });
}

function createMpvClientRuntimeService(): MpvIpcClient {
  const mpvClient = new MpvIpcClient(appState.mpvSocketPath, {
    getResolvedConfig: () => getResolvedConfig(),
    autoStartOverlay: appState.autoStartOverlay,
    setOverlayVisible: (visible) => setOverlayVisible(visible),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      shouldBindVisibleOverlayToMpvSubVisibility(),
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getReconnectTimer: () => appState.reconnectTimer,
    setReconnectTimer: (timer) => {
      appState.reconnectTimer = timer;
    },
    getCurrentSubText: () => appState.currentSubText,
    setCurrentSubText: (text) => {
      appState.currentSubText = text;
    },
    setCurrentSubAssText: (text) => {
      appState.currentSubAssText = text;
    },
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    subtitleWsBroadcast: (text) => {
      subtitleWsService.broadcast(text);
    },
    getOverlayWindowsCount: () => getOverlayWindows().length,
    tokenizeSubtitle: (text) => tokenizeSubtitle(text),
    broadcastToOverlayWindows: (channel, ...channelArgs) => {
      broadcastToOverlayWindows(channel, ...channelArgs);
    },
    updateMpvSubtitleRenderMetrics: (patch) => {
      updateMpvSubtitleRenderMetrics(patch);
    },
    getMpvSubtitleRenderMetrics: () => appState.mpvSubtitleRenderMetrics,
    getPreviousSecondarySubVisibility: () => appState.previousSecondarySubVisibility,
    setPreviousSecondarySubVisibility: (value) => {
      appState.previousSecondarySubVisibility = value;
    },
    showMpvOsd: (text) => {
      showMpvOsd(text);
    },
  });
  bindMpvClientEventHandlers(mpvClient);
  return mpvClient;
}

function updateMpvSubtitleRenderMetrics(
  patch: Partial<MpvSubtitleRenderMetrics>,
): void {
  const { next, changed } = applyMpvSubtitleRenderMetricsPatchService(
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
  return tokenizeSubtitleService(
    text,
    createTokenizerDepsRuntimeService({
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
  return createOverlayWindowService(
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
  if (appState.overlayRuntimeInitialized) {
    return;
  }
  const result = initializeOverlayRuntimeService(
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
        appState.windowTracker = tracker;
      },
      getResolvedConfig: () => getResolvedConfig(),
      getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
      getMpvClient: () => appState.mpvClient,
      getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
      setAnkiIntegration: (integration) => {
        appState.ankiIntegration = integration as AnkiIntegration | null;
      },
      showDesktopNotification,
      createFieldGroupingCallback: () => createFieldGroupingCallback(),
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
  showMpvOsdRuntimeService(
    appState.mpvClient,
    text,
    (line) => {
      console.log(line);
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
    getMpvClient: () => appState.mpvClient,
    getResolvedSubsyncConfig: () => getSubsyncConfig(getResolvedConfig().subsync),
    isSubsyncInProgress: () => appState.subsyncInProgress,
    setSubsyncInProgress: (inProgress: boolean) => {
      appState.subsyncInProgress = inProgress;
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
    openManualPicker: (payload: SubsyncManualPayload) => {
      sendToActiveOverlayWindow("subsync:open-manual", payload, {
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
      subtitleTimingTracker: appState.subtitleTimingTracker,
      writeClipboardText: (text) => clipboard.writeText(text),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

function copyCurrentSubtitle(): void {
  copyCurrentSubtitleService(
    {
      subtitleTimingTracker: appState.subtitleTimingTracker,
      writeClipboardText: (text) => clipboard.writeText(text),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function updateLastCardFromClipboard(): Promise<void> {
  await updateLastCardFromClipboardService(
    {
      ankiIntegration: appState.ankiIntegration,
      readClipboardText: () => clipboard.readText(),
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function triggerFieldGrouping(): Promise<void> {
  await triggerFieldGroupingService(
    {
      ankiIntegration: appState.ankiIntegration,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function markLastCardAsAudioCard(): Promise<void> {
  await markLastCardAsAudioCardService(
    {
      ankiIntegration: appState.ankiIntegration,
      showMpvOsd: (text) => showMpvOsd(text),
    },
  );
}

async function mineSentenceCard(): Promise<void> {
  await mineSentenceCardService(
    {
      ankiIntegration: appState.ankiIntegration,
      mpvClient: appState.mpvClient,
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
      subtitleTimingTracker: appState.subtitleTimingTracker,
      ankiIntegration: appState.ankiIntegration,
      getCurrentSecondarySubText: () =>
        appState.mpvClient?.currentSecondarySubText || undefined,
      showMpvOsd: (text) => showMpvOsd(text),
      logError: (message, err) => {
        console.error(message, err);
      },
    },
  );
}

function registerOverlayShortcuts(): void {
  appState.shortcutsRegistered = registerOverlayShortcutsService(
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
  appState.shortcutsRegistered = unregisterOverlayShortcutsRuntimeService(
    appState.shortcutsRegistered,
    getOverlayShortcutLifecycleDeps(),
  );
}

function shouldOverlayShortcutsBeActive(): boolean { return appState.overlayRuntimeInitialized; }
function syncOverlayShortcuts(): void {
  appState.shortcutsRegistered = syncOverlayShortcutsRuntimeService(
    shouldOverlayShortcutsBeActive(),
    appState.shortcutsRegistered,
    getOverlayShortcutLifecycleDeps(),
  );
}
function refreshOverlayShortcuts(): void {
  appState.shortcutsRegistered = refreshOverlayShortcutsRuntimeService(
    shouldOverlayShortcutsBeActive(),
    appState.shortcutsRegistered,
    getOverlayShortcutLifecycleDeps(),
  );
}

function updateVisibleOverlayVisibility(): void {
  updateVisibleOverlayVisibilityService(
    {
      visibleOverlayVisible: overlayManager.getVisibleOverlayVisible(),
      mainWindow: overlayManager.getMainWindow(),
      windowTracker: appState.windowTracker,
      trackerNotReadyWarningShown: appState.trackerNotReadyWarningShown,
      setTrackerNotReadyWarningShown: (shown) => {
        appState.trackerNotReadyWarningShown = shown;
      },
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
      windowTracker: appState.windowTracker,
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
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    setMpvSubVisibility: (mpvSubVisible) => {
      setMpvSubVisibilityRuntimeService(appState.mpvClient, mpvSubVisible);
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
  const layer = overlayModalAutoShownLayer.get(modal);
  overlayModalAutoShownLayer.delete(modal);
  if (!layer) return;
  const shouldKeepLayerVisible = [...restoreVisibleOverlayOnModalClose].some(
    (pendingModal) => overlayModalAutoShownLayer.get(pendingModal) === layer,
  );
  if (shouldKeepLayerVisible) return;

  if (layer === "visible") {
    const mainWindow = overlayManager.getMainWindow();
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.hide();
    }
    return;
  }
  const invisibleWindow = overlayManager.getInvisibleWindow();
  if (invisibleWindow && !invisibleWindow.isDestroyed()) {
    invisibleWindow.hide();
  }
}

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcService(command, createMpvCommandRuntimeServiceDeps());
}

function createMpvCommandRuntimeServiceDeps(): HandleMpvCommandFromIpcOptions {
  return {
    specialCommands: SPECIAL_COMMANDS,
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    runtimeOptionsCycle: (id, direction) => {
      if (!appState.runtimeOptionsManager) {
        return { ok: false, error: "Runtime options manager unavailable" };
      }
      return applyRuntimeOptionResultRuntimeService(
        appState.runtimeOptionsManager.cycleOption(id, direction),
        (text) => showMpvOsd(text),
      );
    },
    showMpvOsd: (text) => showMpvOsd(text),
    mpvReplaySubtitle: () => replayCurrentSubtitleRuntimeService(appState.mpvClient),
    mpvPlayNextSubtitle: () => playNextSubtitleRuntimeService(appState.mpvClient),
    mpvSendCommand: (rawCommand) =>
      sendMpvCommandRuntimeService(appState.mpvClient, rawCommand),
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
  };
}

async function runSubsyncManualFromIpc(
  request: SubsyncManualRunRequest,
): Promise<SubsyncResult> {
  return runSubsyncManualFromIpcRuntimeService(request, getSubsyncRuntimeDeps());
}

const runtimeOptionsIpcDeps = {
  setRuntimeOption: (id: string, value: unknown) =>
    setRuntimeOptionFromIpcRuntimeService(
      appState.runtimeOptionsManager,
      id as RuntimeOptionId,
      value as RuntimeOptionValue,
      (text) => showMpvOsd(text),
    ),
  cycleRuntimeOption: (id: string, direction: 1 | -1) =>
    cycleRuntimeOptionFromIpcRuntimeService(
      appState.runtimeOptionsManager,
      id as RuntimeOptionId,
      direction,
      (text) => showMpvOsd(text),
    ),
};

registerIpcHandlersService(
  createIpcDepsRuntimeService(createMainIpcRuntimeServiceDeps()),
);

registerAnkiJimakuIpcRuntimeService(
  createAnkiJimakuIpcRuntimeServiceDeps(),
);

function createMainIpcRuntimeServiceDeps(): IpcDepsRuntimeOptions {
  return {
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
    tokenizeCurrentSubtitle: () => tokenizeSubtitle(appState.currentSubText),
    getCurrentSubtitleAss: () => appState.currentSubAssText,
    getMpvSubtitleRenderMetrics: () => appState.mpvSubtitleRenderMetrics,
    getSubtitlePosition: () => loadSubtitlePosition(),
    getSubtitleStyle: () => getResolvedConfig().subtitleStyle ?? null,
    saveSubtitlePosition: (position) =>
      saveSubtitlePosition(position as SubtitlePosition),
    getMecabTokenizer: () => appState.mecabTokenizer,
    handleMpvCommand: (command) => handleMpvCommandFromIpc(command),
    getKeybindings: () => appState.keybindings,
    getSecondarySubMode: () => appState.secondarySubMode,
    getMpvClient: () => appState.mpvClient,
    runSubsyncManual: (request) =>
      runSubsyncManualFromIpc(request as SubsyncManualRunRequest),
    getAnkiConnectStatus: () => appState.ankiIntegration !== null,
    getRuntimeOptions: () => getRuntimeOptionsState(),
    setRuntimeOption: runtimeOptionsIpcDeps.setRuntimeOption,
    cycleRuntimeOption: runtimeOptionsIpcDeps.cycleRuntimeOption,
    reportOverlayContentBounds: (payload) => {
      overlayContentMeasurementStore.report(payload);
    },
  };
}

function createAnkiJimakuIpcRuntimeServiceDeps(): AnkiJimakuIpcRuntimeOptions {
  return {
    patchAnkiConnectEnabled: (enabled) => {
      configService.patchRawConfig({ ankiConnect: { enabled } });
    },
    getResolvedConfig: () => getResolvedConfig(),
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getMpvClient: () => appState.mpvClient,
    getAnkiIntegration: () => appState.ankiIntegration,
    setAnkiIntegration: (integration) => {
      appState.ankiIntegration = integration;
    },
    showDesktopNotification,
    createFieldGroupingCallback: () => createFieldGroupingCallback(),
    broadcastRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    getFieldGroupingResolver: () => getFieldGroupingResolver(),
    setFieldGroupingResolver: (resolver) => setFieldGroupingResolver(resolver),
    parseMediaInfo: (mediaPath) => parseMediaInfo(resolveMediaPathForJimaku(mediaPath)),
    getCurrentMediaPath: () => appState.currentMediaPath,
    jimakuFetchJson: <T>(
      endpoint: string,
      query?: Record<string, string | number | boolean | null | undefined>,
    ): Promise<JimakuApiResponse<T>> => jimakuFetchJson(endpoint, query),
    getJimakuMaxEntryResults: () => getJimakuMaxEntryResults(),
    getJimakuLanguagePreference: () => getJimakuLanguagePreference(),
    resolveJimakuApiKey: () => resolveJimakuApiKey(),
    isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath),
    downloadToFile: (url, destPath, headers) =>
      downloadToFile(url, destPath, headers),
  };
}
