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
  Config,
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
  RuntimeOptionApplyResult,
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
  jimakuFetchJson as jimakuFetchJsonRequest,
  parseMediaInfo,
  resolveJimakuApiKey as resolveJimakuApiKeyFromConfig,
} from "./jimaku/utils";
import {
  getSubsyncConfig,
} from "./subsync/utils";
import {
  CliArgs,
  CliCommandSource,
  commandNeedsOverlayRuntime,
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
  unregisterOverlayShortcutsService,
} from "./core/services/overlay-shortcut-service";
import { runOverlayShortcutLocalFallback } from "./core/services/overlay-shortcut-fallback-runner";
import { createOverlayShortcutRuntimeHandlers } from "./core/services/overlay-shortcut-runtime-service";
import { showDesktopNotification } from "./core/utils/notification";
import { openYomitanSettingsWindow } from "./core/services/yomitan-settings-service";
import { tokenizeSubtitleService } from "./core/services/tokenizer-service";
import { loadYomitanExtensionService } from "./core/services/yomitan-extension-loader-service";
import {
  createOverlayWindowService,
  enforceOverlayLayerOrderService,
  ensureOverlayWindowLevelService,
  updateOverlayBoundsService,
} from "./core/services/overlay-window-service";
import { createFieldGroupingCallbackService } from "./core/services/field-grouping-service";
import { initializeOverlayRuntimeService } from "./core/services/overlay-runtime-init-service";
import {
  setInvisibleOverlayVisibleService,
  setVisibleOverlayVisibleService,
  syncInvisibleOverlayMousePassthroughService,
} from "./core/services/overlay-visibility-runtime-service";
import {
  MpvIpcClient,
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
} from "./core/services/mpv-service";
import {
  handleMpvCommandFromIpcService,
  runSubsyncManualFromIpcService,
} from "./core/services/ipc-command-service";
import { sendToVisibleOverlayService } from "./core/services/overlay-send-service";
import {
  runSubsyncManualService,
  triggerSubsyncFromConfigService,
} from "./core/services/subsync-service";
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
let pendingMultiCopy = false;
let pendingMultiCopyTimeout: ReturnType<typeof setTimeout> | null = null;
let multiCopyDigitShortcuts: string[] = [];
let multiCopyEscapeShortcut: string | null = null;
let pendingMineSentenceMultiple = false;
let pendingMineSentenceMultipleTimeout: ReturnType<typeof setTimeout> | null =
  null;
let overlayRuntimeInitialized = false;
let mineSentenceDigitShortcuts: string[] = [];
let mineSentenceEscapeShortcut: string | null = null;
let fieldGroupingResolver: ((choice: KikuFieldGroupingChoice) => void) | null =
  null;
let runtimeOptionsManager: RuntimeOptionsManager | null = null;
let trackerNotReadyWarningShown = false;
let overlayDebugVisualizationEnabled = false;
type OverlayHostedModal = "runtime-options" | "subsync";
const restoreVisibleOverlayOnModalClose = new Set<OverlayHostedModal>();

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, "subtitle-positions");

interface LoadConfigResult {
  success: boolean;
  config: Config;
}

function loadConfig(): LoadConfigResult {
  const config = configService.getRawConfig();
  return { success: true, config };
}

function saveConfig(config: Config): void {
  try {
    configService.saveRawConfig(config);
    configService.reloadConfig();
  } catch (err) {
    console.error("Failed to save config:", (err as Error).message);
  }
}

function getRuntimeOptionsState(): RuntimeOptionState[] {
  if (!runtimeOptionsManager) return [];
  return runtimeOptionsManager.listOptions();
}

function getOverlayWindows(): BrowserWindow[] {
  const windows: BrowserWindow[] = [];
  if (mainWindow && !mainWindow.isDestroyed()) {
    windows.push(mainWindow);
  }
  if (invisibleWindow && !invisibleWindow.isDestroyed()) {
    windows.push(invisibleWindow);
  }
  return windows;
}

function broadcastToOverlayWindows(channel: string, ...args: unknown[]): void {
  for (const window of getOverlayWindows()) {
    window.webContents.send(channel, ...args);
  }
}

function broadcastRuntimeOptionsChanged(): void {
  broadcastToOverlayWindows(
    "runtime-options:changed",
    getRuntimeOptionsState(),
  );
}

function setOverlayDebugVisualizationEnabled(enabled: boolean): void {
  if (overlayDebugVisualizationEnabled === enabled) return;
  overlayDebugVisualizationEnabled = enabled;
  broadcastToOverlayWindows(
    "overlay-debug-visualization:set",
    overlayDebugVisualizationEnabled,
  );
}

function applyRuntimeOptionResult(
  result: RuntimeOptionApplyResult,
): RuntimeOptionApplyResult {
  if (result.ok && result.osdMessage) {
    showMpvOsd(result.osdMessage);
  }
  return result;
}

function openRuntimeOptionsPalette(): void {
  sendToVisibleOverlay("runtime-options:open", undefined, {
    restoreOnModalClose: "runtime-options",
  });
}

function getResolvedConfig() {
  return configService.getConfig();
}

function getInitialInvisibleOverlayVisibility(): boolean {
  const visibility = getResolvedConfig().invisibleOverlay.startupVisibility;
  if (visibility === "visible") return true;
  if (visibility === "hidden") return false;
  if (process.platform === "linux") return false;
  return true;
}

function shouldAutoInitializeOverlayRuntimeFromConfig(): boolean {
  const config = getResolvedConfig();
  if (config.auto_start_overlay === true) return true;
  if (config.invisibleOverlay.startupVisibility === "visible") return true;
  return false;
}

function shouldBindVisibleOverlayToMpvSubVisibility(): boolean {
  return getResolvedConfig().bind_visible_overlay_to_mpv_sub_visibility;
}

function isAutoUpdateEnabledRuntime(): boolean {
  const value = runtimeOptionsManager?.getOptionValue(
    "anki.autoUpdateNewCards",
  );
  if (typeof value === "boolean") return value;
  const config = getResolvedConfig();
  return config.ankiConnect?.behavior?.autoUpdateNewCards !== false;
}

function getJimakuConfig(): JimakuConfig {
  const config = getResolvedConfig();
  return config.jimaku ?? {};
}

function getJimakuBaseUrl(): string {
  const config = getJimakuConfig();
  return config.apiBaseUrl || DEFAULT_CONFIG.jimaku.apiBaseUrl;
}

function getJimakuLanguagePreference(): JimakuLanguagePreference {
  const config = getJimakuConfig();
  return config.languagePreference || DEFAULT_CONFIG.jimaku.languagePreference;
}

function getJimakuMaxEntryResults(): number {
  const config = getJimakuConfig();
  const value = config.maxEntryResults;
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  return DEFAULT_CONFIG.jimaku.maxEntryResults;
}

async function resolveJimakuApiKey(): Promise<string | null> {
  return resolveJimakuApiKeyFromConfig(getJimakuConfig());
}

async function jimakuFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined> = {},
): Promise<JimakuApiResponse<T>> {
  const apiKey = await resolveJimakuApiKey();
  if (!apiKey) {
    return {
      ok: false,
      error: {
        error:
          "Jimaku API key not set. Configure jimaku.apiKey or jimaku.apiKeyCommand.",
        code: 401,
      },
    };
  }

  return jimakuFetchJsonRequest<T>(endpoint, query, {
    baseUrl: getJimakuBaseUrl(),
    apiKey,
  });
}

function getSubtitlePositionFilePath(mediaPath: string): string {
  const key = normalizeMediaPathForSubtitlePosition(mediaPath);
  const hash = crypto.createHash("sha256").update(key).digest("hex");
  return path.join(SUBTITLE_POSITIONS_DIR, `${hash}.json`);
}

function normalizeMediaPathForSubtitlePosition(mediaPath: string): string {
  const trimmed = mediaPath.trim();
  if (!trimmed) return trimmed;

  if (
    /^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(trimmed) ||
    /^ytsearch:/.test(trimmed)
  ) {
    return trimmed;
  }

  const resolved = path.resolve(trimmed);
  let normalized = resolved;
  try {
    if (fs.existsSync(resolved)) {
      normalized = fs.realpathSync(resolved);
    }
  } catch {
    normalized = resolved;
  }

  if (process.platform === "win32") {
    normalized = normalized.toLowerCase();
  }

  return normalized;
}

function persistSubtitlePosition(position: SubtitlePosition): void {
  if (!currentMediaPath) return;
  if (!fs.existsSync(SUBTITLE_POSITIONS_DIR)) {
    fs.mkdirSync(SUBTITLE_POSITIONS_DIR, { recursive: true });
  }
  const positionPath = getSubtitlePositionFilePath(currentMediaPath);
  fs.writeFileSync(positionPath, JSON.stringify(position, null, 2));
}

function loadSubtitlePosition(): SubtitlePosition | null {
  const fallbackPosition = getResolvedConfig().subtitlePosition;
  if (!currentMediaPath) {
    subtitlePosition = fallbackPosition;
    return subtitlePosition;
  }

  try {
    const positionPath = getSubtitlePositionFilePath(currentMediaPath);
    if (!fs.existsSync(positionPath)) {
      subtitlePosition = fallbackPosition;
      return subtitlePosition;
    }

    const data = fs.readFileSync(positionPath, "utf-8");
    const parsed = JSON.parse(data) as Partial<SubtitlePosition>;
    if (
      parsed &&
      typeof parsed.yPercent === "number" &&
      Number.isFinite(parsed.yPercent)
    ) {
      subtitlePosition = { yPercent: parsed.yPercent };
    } else {
      subtitlePosition = fallbackPosition;
    }
  } catch (err) {
    console.error("Failed to load subtitle position:", (err as Error).message);
    subtitlePosition = fallbackPosition;
  }

  return subtitlePosition;
}

function saveSubtitlePosition(position: SubtitlePosition): void {
  subtitlePosition = position;
  if (!currentMediaPath) {
    pendingSubtitlePosition = position;
    console.warn("Queued subtitle position save - no media path yet");
    return;
  }

  try {
    persistSubtitlePosition(position);
    pendingSubtitlePosition = null;
  } catch (err) {
    console.error("Failed to save subtitle position:", (err as Error).message);
  }
}

function updateCurrentMediaPath(mediaPath: unknown): void {
  const nextPath =
    typeof mediaPath === "string" && mediaPath.trim().length > 0
      ? mediaPath
      : null;
  if (nextPath === currentMediaPath) return;
  currentMediaPath = nextPath;

  if (currentMediaPath && pendingSubtitlePosition) {
    try {
      persistSubtitlePosition(pendingSubtitlePosition);
      subtitlePosition = pendingSubtitlePosition;
      pendingSubtitlePosition = null;
    } catch (err) {
      console.error(
        "Failed to persist queued subtitle position:",
        (err as Error).message,
      );
    }
  }

  const position = loadSubtitlePosition();
  broadcastToOverlayWindows("subtitle-position:set", position);
}

const AUTOSUBSYNC_SPINNER_FRAMES = ["|", "/", "-", "\\"];
let subsyncInProgress = false;

async function runWithSubsyncSpinner<T>(
  task: () => Promise<T>,
  label = "Subsync: syncing",
): Promise<T> {
  let frame = 0;
  showMpvOsd(`${label} ${AUTOSUBSYNC_SPINNER_FRAMES[0]}`);
  const timer = setInterval(() => {
    frame = (frame + 1) % AUTOSUBSYNC_SPINNER_FRAMES.length;
    showMpvOsd(`${label} ${AUTOSUBSYNC_SPINNER_FRAMES[frame]}`);
  }, 150);

  try {
    return await task();
  } finally {
    clearInterval(timer);
  }
}

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

if (initialArgs.generateConfig && !shouldStartApp(initialArgs)) {
  generateDefaultConfigFile(initialArgs, {
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
} else {
  const gotTheLock = app.requestSingleInstanceLock();

  if (!gotTheLock) {
    app.quit();
  } else {
    app.on("second-instance", (_event, argv) => {
      handleCliCommand(parseArgs(argv), "second-instance");
    });
    if (initialArgs.help && !shouldStartApp(initialArgs)) {
      printHelp(DEFAULT_TEXTHOOKER_PORT);
      app.quit();
    } else if (!shouldStartApp(initialArgs)) {
      if (initialArgs.stop && !initialArgs.start) {
        app.quit();
      } else {
        console.error("No running instance. Use --start to launch the app.");
        app.quit();
      }
    } else {
      app.whenReady().then(async () => {
        loadSubtitlePosition();
        keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);

        mpvClient = new MpvIpcClient(mpvSocketPath, {
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
        });

        configService.reloadConfig();
        const config = getResolvedConfig();
        for (const warning of configService.getWarnings()) {
          console.warn(
            `[config] ${warning.path}: ${warning.message} value=${JSON.stringify(warning.value)} fallback=${JSON.stringify(warning.fallback)}`,
          );
        }
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
        secondarySubMode = config.secondarySub?.defaultMode ?? "hover";
        const wsConfig = config.websocket || {};
        const wsEnabled = wsConfig.enabled ?? "auto";
        const wsPort = wsConfig.port || DEFAULT_CONFIG.websocket.port;

        if (
          wsEnabled === true ||
          (wsEnabled === "auto" && !hasMpvWebsocketPlugin())
        ) {
          subtitleWsService.start(wsPort, () => currentSubText);
        } else if (wsEnabled === "auto") {
          console.log(
            "mpv_websocket detected, skipping built-in WebSocket server",
          );
        }

        mecabTokenizer = new MecabTokenizer();
        await mecabTokenizer.checkAvailability();

        subtitleTimingTracker = new SubtitleTimingTracker();

        await loadYomitanExtension();
        if (texthookerOnlyMode) {
          console.log("Texthooker-only mode enabled; skipping overlay window.");
        } else if (shouldAutoInitializeOverlayRuntimeFromConfig()) {
          initializeOverlayRuntime();
        } else {
          console.log(
            "Overlay runtime deferred: waiting for explicit overlay command.",
          );
        }

        handleInitialArgs();
      });

      app.on("window-all-closed", () => {
        if (process.platform !== "darwin") {
          app.quit();
        }
      });

      app.on("will-quit", () => {
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
      });

      app.on("activate", () => {
        if (
          overlayRuntimeInitialized &&
          BrowserWindow.getAllWindows().length === 0
        ) {
          createMainWindow();
          createInvisibleWindow();
          updateVisibleOverlayVisibility();
          updateInvisibleOverlayVisibility();
        }
      });
    }
  }
}

function handleCliCommand(
  args: CliArgs,
  source: CliCommandSource = "initial",
): void {
  const hasNonStartAction =
    args.stop ||
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.toggleInvisibleOverlay ||
    args.settings ||
    args.show ||
    args.hide ||
    args.showVisibleOverlay ||
    args.hideVisibleOverlay ||
    args.showInvisibleOverlay ||
    args.hideInvisibleOverlay ||
    args.copySubtitle ||
    args.copySubtitleMultiple ||
    args.mineSentence ||
    args.mineSentenceMultiple ||
    args.updateLastCardFromClipboard ||
    args.toggleSecondarySub ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.openRuntimeOptions ||
    args.texthooker ||
    args.help;
  const ignoreStart = source === "second-instance" && args.start;
  if (ignoreStart && !hasNonStartAction) {
    console.log("Ignoring --start because SubMiner is already running.");
    return;
  }
  const shouldStart =
    !ignoreStart &&
    (args.start ||
      (source === "initial" &&
        (args.toggle ||
          args.toggleVisibleOverlay ||
          args.toggleInvisibleOverlay)));
  const needsOverlayRuntime = commandNeedsOverlayRuntime(args);

  if (args.socketPath !== undefined) {
    mpvSocketPath = args.socketPath;
    if (mpvClient) {
      mpvClient.setSocketPath(mpvSocketPath);
    }
  }
  if (args.texthookerPort !== undefined) {
    if (texthookerService.isRunning()) {
      console.warn(
        "Ignoring --port override because the texthooker server is already running.",
      );
    } else {
      texthookerPort = args.texthookerPort;
    }
  }

  if (args.stop) {
    console.log("Stopping SubMiner...");
    app.quit();
    return;
  }

  if (needsOverlayRuntime && !overlayRuntimeInitialized) {
    initializeOverlayRuntime();
  }

  if (shouldStart && mpvClient) {
    mpvClient.setSocketPath(mpvSocketPath);
    mpvClient.connect();
    console.log(`Starting MPV IPC connection on socket: ${mpvSocketPath}`);
  }

  if (args.toggle || args.toggleVisibleOverlay) {
    toggleVisibleOverlay();
  } else if (args.toggleInvisibleOverlay) {
    toggleInvisibleOverlay();
  } else if (args.settings) {
    setTimeout(() => {
      openYomitanSettings();
    }, 1000);
  } else if (args.show || args.showVisibleOverlay) {
    setVisibleOverlayVisible(true);
  } else if (args.hide || args.hideVisibleOverlay) {
    setVisibleOverlayVisible(false);
  } else if (args.showInvisibleOverlay) {
    setInvisibleOverlayVisible(true);
  } else if (args.hideInvisibleOverlay) {
    setInvisibleOverlayVisible(false);
  } else if (args.copySubtitle) {
    copyCurrentSubtitle();
  } else if (args.copySubtitleMultiple) {
    startPendingMultiCopy(getConfiguredShortcuts().multiCopyTimeoutMs);
  } else if (args.mineSentence) {
    mineSentenceCard().catch((err) => {
      console.error("mineSentenceCard failed:", err);
      showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
    });
  } else if (args.mineSentenceMultiple) {
    startPendingMineSentenceMultiple(getConfiguredShortcuts().multiCopyTimeoutMs);
  } else if (args.updateLastCardFromClipboard) {
    updateLastCardFromClipboard().catch((err) => {
      console.error("updateLastCardFromClipboard failed:", err);
      showMpvOsd(`Update failed: ${(err as Error).message}`);
    });
  } else if (args.toggleSecondarySub) {
    cycleSecondarySubMode();
  } else if (args.triggerFieldGrouping) {
    triggerFieldGrouping().catch((err) => {
      console.error("triggerFieldGrouping failed:", err);
      showMpvOsd(`Field grouping failed: ${(err as Error).message}`);
    });
  } else if (args.triggerSubsync) {
    triggerSubsyncFromConfig().catch((err) => {
      console.error("triggerSubsyncFromConfig failed:", err);
      showMpvOsd(`Subsync failed: ${(err as Error).message}`);
    });
  } else if (args.markAudioCard) {
    markLastCardAsAudioCard().catch((err) => {
      console.error("markLastCardAsAudioCard failed:", err);
      showMpvOsd(`Audio card failed: ${(err as Error).message}`);
    });
  } else if (args.openRuntimeOptions) {
    openRuntimeOptionsPalette();
  } else if (args.texthooker) {
    if (!texthookerService.isRunning()) {
      texthookerService.start(texthookerPort);
    }
    const config = getResolvedConfig();
    const openBrowser = config.texthooker?.openBrowser !== false;
    if (openBrowser) {
      shell.openExternal(`http://127.0.0.1:${texthookerPort}`);
    }
    console.log(`Texthooker available at http://127.0.0.1:${texthookerPort}`);
  } else if (args.help) {
    printHelp(DEFAULT_TEXTHOOKER_PORT);
    if (!mainWindow) app.quit();
  }
}

function handleInitialArgs(): void {
  handleCliCommand(initialArgs, "initial");
}

function updateMpvSubtitleRenderMetrics(
  patch: Partial<MpvSubtitleRenderMetrics>,
): void {
  const patchOsd = patch.osdDimensions;
  const nextOsdDimensions =
    patchOsd &&
    typeof patchOsd.w === "number" &&
    typeof patchOsd.h === "number" &&
    typeof patchOsd.ml === "number" &&
    typeof patchOsd.mr === "number" &&
    typeof patchOsd.mt === "number" &&
    typeof patchOsd.mb === "number"
      ? {
          w: asFiniteNumber(patchOsd.w, 0, 1, 100000),
          h: asFiniteNumber(patchOsd.h, 0, 1, 100000),
          ml: asFiniteNumber(patchOsd.ml, 0, 0, 100000),
          mr: asFiniteNumber(patchOsd.mr, 0, 0, 100000),
          mt: asFiniteNumber(patchOsd.mt, 0, 0, 100000),
          mb: asFiniteNumber(patchOsd.mb, 0, 0, 100000),
        }
      : patchOsd === null
        ? null
        : mpvSubtitleRenderMetrics.osdDimensions;

  const next: MpvSubtitleRenderMetrics = {
    subPos: asFiniteNumber(
      patch.subPos,
      mpvSubtitleRenderMetrics.subPos,
      0,
      150,
    ),
    subFontSize: asFiniteNumber(
      patch.subFontSize,
      mpvSubtitleRenderMetrics.subFontSize,
      1,
      200,
    ),
    subScale: asFiniteNumber(
      patch.subScale,
      mpvSubtitleRenderMetrics.subScale,
      0.1,
      10,
    ),
    subMarginY: asFiniteNumber(
      patch.subMarginY,
      mpvSubtitleRenderMetrics.subMarginY,
      0,
      200,
    ),
    subMarginX: asFiniteNumber(
      patch.subMarginX,
      mpvSubtitleRenderMetrics.subMarginX,
      0,
      200,
    ),
    subFont: asString(patch.subFont, mpvSubtitleRenderMetrics.subFont),
    subSpacing: asFiniteNumber(
      patch.subSpacing,
      mpvSubtitleRenderMetrics.subSpacing,
      -100,
      100,
    ),
    subBold: asBoolean(patch.subBold, mpvSubtitleRenderMetrics.subBold),
    subItalic: asBoolean(patch.subItalic, mpvSubtitleRenderMetrics.subItalic),
    subBorderSize: asFiniteNumber(
      patch.subBorderSize,
      mpvSubtitleRenderMetrics.subBorderSize,
      0,
      100,
    ),
    subShadowOffset: asFiniteNumber(
      patch.subShadowOffset,
      mpvSubtitleRenderMetrics.subShadowOffset,
      0,
      100,
    ),
    subAssOverride: asString(
      patch.subAssOverride,
      mpvSubtitleRenderMetrics.subAssOverride,
    ),
    subScaleByWindow: asBoolean(
      patch.subScaleByWindow,
      mpvSubtitleRenderMetrics.subScaleByWindow,
    ),
    subUseMargins: asBoolean(
      patch.subUseMargins,
      mpvSubtitleRenderMetrics.subUseMargins,
    ),
    osdHeight: asFiniteNumber(
      patch.osdHeight,
      mpvSubtitleRenderMetrics.osdHeight,
      1,
      10000,
    ),
    osdDimensions: nextOsdDimensions,
  };

  const changed =
    next.subPos !== mpvSubtitleRenderMetrics.subPos ||
    next.subFontSize !== mpvSubtitleRenderMetrics.subFontSize ||
    next.subScale !== mpvSubtitleRenderMetrics.subScale ||
    next.subMarginY !== mpvSubtitleRenderMetrics.subMarginY ||
    next.subMarginX !== mpvSubtitleRenderMetrics.subMarginX ||
    next.subFont !== mpvSubtitleRenderMetrics.subFont ||
    next.subSpacing !== mpvSubtitleRenderMetrics.subSpacing ||
    next.subBold !== mpvSubtitleRenderMetrics.subBold ||
    next.subItalic !== mpvSubtitleRenderMetrics.subItalic ||
    next.subBorderSize !== mpvSubtitleRenderMetrics.subBorderSize ||
    next.subShadowOffset !== mpvSubtitleRenderMetrics.subShadowOffset ||
    next.subAssOverride !== mpvSubtitleRenderMetrics.subAssOverride ||
    next.subScaleByWindow !== mpvSubtitleRenderMetrics.subScaleByWindow ||
    next.subUseMargins !== mpvSubtitleRenderMetrics.subUseMargins ||
    next.osdHeight !== mpvSubtitleRenderMetrics.osdHeight ||
    JSON.stringify(next.osdDimensions) !==
      JSON.stringify(mpvSubtitleRenderMetrics.osdDimensions);

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
  const now = Date.now();
  if (now - lastSecondarySubToggleAtMs < 120) {
    return;
  }
  lastSecondarySubToggleAtMs = now;

  const cycle: SecondarySubMode[] = ["hidden", "visible", "hover"];
  const idx = cycle.indexOf(secondarySubMode);
  secondarySubMode = cycle[(idx + 1) % cycle.length];
  broadcastToOverlayWindows("secondary-subtitle:mode", secondarySubMode);
  showMpvOsd(`Secondary subtitle: ${secondarySubMode}`);
}

function showMpvOsd(text: string): void {
  if (mpvClient && mpvClient.connected && mpvClient.send) {
    mpvClient.send({
      command: ["show-text", text, "3000"],
    });
  } else {
    console.log("OSD (MPV not connected):", text);
  }
}

function getSubsyncServiceDeps() {
  return {
    getMpvClient: () => mpvClient,
    getResolvedConfig: () => getSubsyncConfig(getResolvedConfig().subsync),
    isSubsyncInProgress: () => subsyncInProgress,
    setSubsyncInProgress: (inProgress: boolean) => {
      subsyncInProgress = inProgress;
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
    runWithSubsyncSpinner: <T>(task: () => Promise<T>) =>
      runWithSubsyncSpinner(task),
    openManualPicker: (payload: SubsyncManualPayload) => {
      sendToVisibleOverlay("subsync:open-manual", payload, {
        restoreOnModalClose: "subsync",
      });
    },
  };
}

async function triggerSubsyncFromConfig(): Promise<void> {
  await triggerSubsyncFromConfigService(getSubsyncServiceDeps());
}

function cancelPendingMultiCopy(): void {
  if (!pendingMultiCopy) return;

  pendingMultiCopy = false;
  if (pendingMultiCopyTimeout) {
    clearTimeout(pendingMultiCopyTimeout);
    pendingMultiCopyTimeout = null;
  }

  for (const shortcut of multiCopyDigitShortcuts) {
    globalShortcut.unregister(shortcut);
  }
  multiCopyDigitShortcuts = [];

  if (multiCopyEscapeShortcut) {
    globalShortcut.unregister(multiCopyEscapeShortcut);
    multiCopyEscapeShortcut = null;
  }
}

function startPendingMultiCopy(timeoutMs: number): void {
  cancelPendingMultiCopy();
  pendingMultiCopy = true;

  for (let i = 1; i <= 9; i++) {
    const shortcut = i.toString();
    if (
      globalShortcut.register(shortcut, () => {
        handleMultiCopyDigit(i);
      })
    ) {
      multiCopyDigitShortcuts.push(shortcut);
    }
  }

  if (
    globalShortcut.register("Escape", () => {
      cancelPendingMultiCopy();
      showMpvOsd("Cancelled");
    })
  ) {
    multiCopyEscapeShortcut = "Escape";
  }

  pendingMultiCopyTimeout = setTimeout(() => {
    cancelPendingMultiCopy();
    showMpvOsd("Copy timeout");
  }, timeoutMs);

  showMpvOsd("Copy how many lines? Press 1-9 (Esc to cancel)");
}

function handleMultiCopyDigit(count: number): void {
  if (!pendingMultiCopy || !subtitleTimingTracker) return;

  cancelPendingMultiCopy();

  const availableCount = Math.min(count, 200); // Max history size
  const blocks = subtitleTimingTracker.getRecentBlocks(availableCount);

  if (blocks.length === 0) {
    showMpvOsd("No subtitle history available");
    return;
  }

  const actualCount = blocks.length;
  const clipboardText = blocks.join("\n\n");
  clipboard.writeText(clipboardText);

  if (actualCount < count) {
    showMpvOsd(`Only ${actualCount} lines available, copied ${actualCount}`);
  } else {
    showMpvOsd(`Copied ${actualCount} lines`);
  }
}

function copyCurrentSubtitle(): void {
  if (!subtitleTimingTracker) {
    showMpvOsd("Subtitle tracker not available");
    return;
  }

  const currentSubtitle = subtitleTimingTracker.getCurrentSubtitle();
  if (!currentSubtitle) {
    showMpvOsd("No current subtitle");
    return;
  }

  clipboard.writeText(currentSubtitle);
  showMpvOsd("Copied subtitle");
}

async function updateLastCardFromClipboard(): Promise<void> {
  if (!ankiIntegration) {
    showMpvOsd("AnkiConnect integration not enabled");
    return;
  }

  const clipboardText = clipboard.readText();
  await ankiIntegration.updateLastAddedFromClipboard(clipboardText);
}

async function triggerFieldGrouping(): Promise<void> {
  if (!ankiIntegration) {
    showMpvOsd("AnkiConnect integration not enabled");
    return;
  }
  await ankiIntegration.triggerFieldGroupingForLastAddedCard();
}

async function markLastCardAsAudioCard(): Promise<void> {
  if (!ankiIntegration) {
    showMpvOsd("AnkiConnect integration not enabled");
    return;
  }
  await ankiIntegration.markLastCardAsAudioCard();
}

async function mineSentenceCard(): Promise<void> {
  if (!ankiIntegration) {
    showMpvOsd("AnkiConnect integration not enabled");
    return;
  }

  if (!mpvClient || !mpvClient.connected) {
    showMpvOsd("MPV not connected");
    return;
  }

  const text = mpvClient.currentSubText;
  if (!text) {
    showMpvOsd("No current subtitle");
    return;
  }

  const startTime = mpvClient.currentSubStart;
  const endTime = mpvClient.currentSubEnd;
  const secondarySub = mpvClient.currentSecondarySubText || undefined;

  await ankiIntegration.createSentenceCard(
    text,
    startTime,
    endTime,
    secondarySub,
  );
}

function cancelPendingMineSentenceMultiple(): void {
  if (!pendingMineSentenceMultiple) return;

  pendingMineSentenceMultiple = false;
  if (pendingMineSentenceMultipleTimeout) {
    clearTimeout(pendingMineSentenceMultipleTimeout);
    pendingMineSentenceMultipleTimeout = null;
  }

  for (const shortcut of mineSentenceDigitShortcuts) {
    globalShortcut.unregister(shortcut);
  }
  mineSentenceDigitShortcuts = [];

  if (mineSentenceEscapeShortcut) {
    globalShortcut.unregister(mineSentenceEscapeShortcut);
    mineSentenceEscapeShortcut = null;
  }
}

function startPendingMineSentenceMultiple(timeoutMs: number): void {
  cancelPendingMineSentenceMultiple();
  pendingMineSentenceMultiple = true;

  for (let i = 1; i <= 9; i++) {
    const shortcut = i.toString();
    if (
      globalShortcut.register(shortcut, () => {
        handleMineSentenceDigit(i);
      })
    ) {
      mineSentenceDigitShortcuts.push(shortcut);
    }
  }

  if (
    globalShortcut.register("Escape", () => {
      cancelPendingMineSentenceMultiple();
      showMpvOsd("Cancelled");
    })
  ) {
    mineSentenceEscapeShortcut = "Escape";
  }

  pendingMineSentenceMultipleTimeout = setTimeout(() => {
    cancelPendingMineSentenceMultiple();
    showMpvOsd("Mine sentence timeout");
  }, timeoutMs);

  showMpvOsd("Mine how many lines? Press 1-9 (Esc to cancel)");
}

function handleMineSentenceDigit(count: number): void {
  if (
    !pendingMineSentenceMultiple ||
    !subtitleTimingTracker ||
    !ankiIntegration
  )
    return;

  cancelPendingMineSentenceMultiple();

  const blocks = subtitleTimingTracker.getRecentBlocks(count);

  if (blocks.length === 0) {
    showMpvOsd("No subtitle history available");
    return;
  }

  const timings: { startTime: number; endTime: number }[] = [];
  for (const block of blocks) {
    const timing = subtitleTimingTracker.findTiming(block);
    if (timing) {
      timings.push(timing);
    }
  }

  if (timings.length === 0) {
    showMpvOsd("Subtitle timing not found");
    return;
  }

  const rangeStart = Math.min(...timings.map((t) => t.startTime));
  const rangeEnd = Math.max(...timings.map((t) => t.endTime));
  const sentence = blocks.join(" ");

  const secondarySub = mpvClient?.currentSecondarySubText || undefined;
  ankiIntegration
    .createSentenceCard(sentence, rangeStart, rangeEnd, secondarySub)
    .catch((err) => {
      console.error("mineSentenceMultiple failed:", err);
      showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
    });
}

function registerOverlayShortcuts(): void {
  const shortcuts = getConfiguredShortcuts();
  const handlers = getOverlayShortcutRuntimeHandlers();
  shortcutsRegistered = registerOverlayShortcutsService(
    shortcuts,
    handlers.overlayHandlers,
  );
}

function unregisterOverlayShortcuts(): void {
  if (!shortcutsRegistered) return;

  cancelPendingMultiCopy();
  cancelPendingMineSentenceMultiple();

  unregisterOverlayShortcutsService(getConfiguredShortcuts());

  shortcutsRegistered = false;
}

function shouldOverlayShortcutsBeActive(): boolean { return overlayRuntimeInitialized; }
function syncOverlayShortcuts(): void { if (shouldOverlayShortcutsBeActive()) { registerOverlayShortcuts(); } else { unregisterOverlayShortcuts(); } }
function refreshOverlayShortcuts(): void { unregisterOverlayShortcuts(); syncOverlayShortcuts(); }

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
  setVisibleOverlayVisibleService({
    visible,
    setVisibleOverlayVisibleState: (nextVisible) => {
      visibleOverlayVisible = nextVisible;
    },
    updateVisibleOverlayVisibility: () => updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () => updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      syncInvisibleOverlayMousePassthrough(),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      shouldBindVisibleOverlayToMpvSubVisibility(),
    isMpvConnected: () => Boolean(mpvClient && mpvClient.connected),
    setMpvSubVisibility: (mpvSubVisible) => {
      if (mpvClient) {
        mpvClient.setSubVisibility(mpvSubVisible);
      }
    },
  });
}

function setInvisibleOverlayVisible(visible: boolean): void {
  setInvisibleOverlayVisibleService({
    visible,
    setInvisibleOverlayVisibleState: (nextVisible) => {
      invisibleOverlayVisible = nextVisible;
    },
    updateInvisibleOverlayVisibility: () => updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      syncInvisibleOverlayMousePassthrough(),
  });
}

function toggleVisibleOverlay(): void { setVisibleOverlayVisible(!visibleOverlayVisible); }
function toggleInvisibleOverlay(): void { setInvisibleOverlayVisible(!invisibleOverlayVisible); }
function setOverlayVisible(visible: boolean): void { setVisibleOverlayVisible(visible); }
function toggleOverlay(): void { toggleVisibleOverlay(); }
function handleOverlayModalClosed(modal: OverlayHostedModal): void { if (!restoreVisibleOverlayOnModalClose.has(modal)) return; restoreVisibleOverlayOnModalClose.delete(modal); if (restoreVisibleOverlayOnModalClose.size === 0) { setVisibleOverlayVisible(false); } }

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcService(command, {
    specialCommands: SPECIAL_COMMANDS,
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    runtimeOptionsCycle: (id, direction) => {
      if (!runtimeOptionsManager) {
        return { ok: false, error: "Runtime options manager unavailable" };
      }
      return applyRuntimeOptionResult(
        runtimeOptionsManager.cycleOption(id, direction),
      );
    },
    showMpvOsd: (text) => showMpvOsd(text),
    mpvReplaySubtitle: () => {
      if (mpvClient) mpvClient.replayCurrentSubtitle();
    },
    mpvPlayNextSubtitle: () => {
      if (mpvClient) mpvClient.playNextSubtitle();
    },
    mpvSendCommand: (rawCommand) => {
      if (!mpvClient) return;
      mpvClient.send({ command: rawCommand });
    },
    isMpvConnected: () => Boolean(mpvClient && mpvClient.connected),
    hasRuntimeOptionsManager: () => runtimeOptionsManager !== null,
  });
}

async function runSubsyncManualFromIpc(
  request: SubsyncManualRunRequest,
): Promise<SubsyncResult> {
  const deps = getSubsyncServiceDeps();
  return runSubsyncManualFromIpcService(request, {
    isSubsyncInProgress: deps.isSubsyncInProgress,
    setSubsyncInProgress: deps.setSubsyncInProgress,
    showMpvOsd: deps.showMpvOsd,
    runWithSpinner: deps.runWithSubsyncSpinner,
    runSubsyncManual: (subsyncRequest) =>
      runSubsyncManualService(subsyncRequest, deps),
  });
}

registerIpcHandlersService({
  getInvisibleWindow: () => invisibleWindow,
  isVisibleOverlayVisible: () => visibleOverlayVisible,
  setInvisibleIgnoreMouseEvents: (ignore, options) => {
    if (!invisibleWindow || invisibleWindow.isDestroyed()) return;
    invisibleWindow.setIgnoreMouseEvents(ignore, options);
  },
  onOverlayModalClosed: (modal) =>
    handleOverlayModalClosed(modal as OverlayHostedModal),
  openYomitanSettings: () => openYomitanSettings(),
  quitApp: () => app.quit(),
  toggleDevTools: () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.toggleDevTools();
    }
  },
  getVisibleOverlayVisibility: () => visibleOverlayVisible,
  toggleVisibleOverlay: () => toggleVisibleOverlay(),
  getInvisibleOverlayVisibility: () => invisibleOverlayVisible,
  tokenizeCurrentSubtitle: () => tokenizeSubtitle(currentSubText),
  getCurrentSubtitleAss: () => currentSubAssText,
  getMpvSubtitleRenderMetrics: () => mpvSubtitleRenderMetrics,
  getSubtitlePosition: () => loadSubtitlePosition(),
  getSubtitleStyle: () => getResolvedConfig().subtitleStyle ?? null,
  saveSubtitlePosition: (position) => saveSubtitlePosition(position as SubtitlePosition),
  getMecabStatus: () =>
    mecabTokenizer
      ? mecabTokenizer.getStatus()
      : { available: false, enabled: false, path: null },
  setMecabEnabled: (enabled) => {
    if (mecabTokenizer) mecabTokenizer.setEnabled(enabled);
  },
  handleMpvCommand: (command) => handleMpvCommandFromIpc(command),
  getKeybindings: () => keybindings,
  getSecondarySubMode: () => secondarySubMode,
  getCurrentSecondarySub: () => mpvClient?.currentSecondarySubText || "",
  runSubsyncManual: (request) =>
    runSubsyncManualFromIpc(request as SubsyncManualRunRequest),
  getAnkiConnectStatus: () => ankiIntegration !== null,
  getRuntimeOptions: () => getRuntimeOptionsState(),
  setRuntimeOption: (id, value) => {
    if (!runtimeOptionsManager) {
      return { ok: false, error: "Runtime options manager unavailable" };
    }
    const result = applyRuntimeOptionResult(
      runtimeOptionsManager.setOptionValue(id as RuntimeOptionId, value as RuntimeOptionValue),
    );
    if (!result.ok && result.error) {
      showMpvOsd(result.error);
    }
    return result;
  },
  cycleRuntimeOption: (id, direction) => {
    if (!runtimeOptionsManager) {
      return { ok: false, error: "Runtime options manager unavailable" };
    }
    const result = applyRuntimeOptionResult(
      runtimeOptionsManager.cycleOption(id as RuntimeOptionId, direction),
    );
    if (!result.ok && result.error) {
      showMpvOsd(result.error);
    }
    return result;
  },
});

/**
 * Create and show a desktop notification with robust icon handling.
 * Supports both file paths (preferred on Linux/Wayland) and data URLs (fallback).
 */
function createFieldGroupingCallback() {
  return createFieldGroupingCallbackService({
    getVisibleOverlayVisible: () => visibleOverlayVisible,
    getInvisibleOverlayVisible: () => invisibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    setInvisibleOverlayVisible: (visible) => setInvisibleOverlayVisible(visible),
    getResolver: () => fieldGroupingResolver,
    setResolver: (resolver) => {
      fieldGroupingResolver = resolver;
    },
    sendRequestToVisibleOverlay: (data) =>
      sendToVisibleOverlay("kiku:field-grouping-request", data),
  });
}

function sendToVisibleOverlay(channel: string, payload?: unknown, options?: { restoreOnModalClose?: OverlayHostedModal }): boolean { return sendToVisibleOverlayService({ mainWindow, visibleOverlayVisible, setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible), channel, payload, restoreOnModalClose: options?.restoreOnModalClose, addRestoreFlag: (modal) => restoreVisibleOverlayOnModalClose.add(modal as OverlayHostedModal) }); }

registerAnkiJimakuIpcRuntimeService({
  patchAnkiConnectEnabled: (enabled) => {
    configService.patchRawConfig({
      ankiConnect: {
        enabled,
      },
    });
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
  getFieldGroupingResolver: () => fieldGroupingResolver,
  setFieldGroupingResolver: (resolver) => {
    fieldGroupingResolver = resolver;
  },
  parseMediaInfo: (mediaPath) => parseMediaInfo(mediaPath),
  getCurrentMediaPath: () => currentMediaPath,
  jimakuFetchJson: (endpoint, query) => jimakuFetchJson(endpoint, query),
  getJimakuMaxEntryResults: () => getJimakuMaxEntryResults(),
  getJimakuLanguagePreference: () => getJimakuLanguagePreference(),
  resolveJimakuApiKey: () => resolveJimakuApiKey(),
  isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath),
  downloadToFile: (url, destPath, headers) => downloadToFile(url, destPath, headers),
});
