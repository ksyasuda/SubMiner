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
  session,
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
import * as net from "net";
import * as http from "http";
import * as https from "https";
import * as os from "os";
import * as fs from "fs";
import * as crypto from "crypto";
import { MecabTokenizer } from "./mecab-tokenizer";
import { mergeTokens } from "./token-merger";
import { createWindowTracker, BaseWindowTracker } from "./window-trackers";
import {
  Config,
  PartOfSpeech,
  MergedToken,
  JimakuApiResponse,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuMediaInfo,
  JimakuConfig,
  JimakuLanguagePreference,
  SubtitleData,
  SubtitlePosition,
  Keybinding,
  WindowGeometry,
  SecondarySubMode,
  MpvClient,
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncResult,
  KikuFieldGroupingRequestData,
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
  sortJimakuFiles,
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
import { showDesktopNotification } from "./core/utils/notification";
import { openYomitanSettingsWindow } from "./core/services/yomitan-settings-service";
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
import { registerAnkiJimakuIpcHandlers } from "./core/services/anki-jimaku-ipc-service";
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

        mpvClient = new MpvIpcClient(mpvSocketPath);

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

interface MpvMessage {
  event?: string;
  name?: string;
  data?: unknown;
  request_id?: number;
  error?: string;
}

const MPV_REQUEST_ID_SUBTEXT = 101;
const MPV_REQUEST_ID_PATH = 102;
const MPV_REQUEST_ID_SECONDARY_SUBTEXT = 103;
const MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY = 104;
const MPV_REQUEST_ID_AID = 105;
const MPV_REQUEST_ID_SUB_POS = 106;
const MPV_REQUEST_ID_SUB_FONT_SIZE = 107;
const MPV_REQUEST_ID_SUB_SCALE = 108;
const MPV_REQUEST_ID_SUB_MARGIN_Y = 109;
const MPV_REQUEST_ID_SUB_MARGIN_X = 110;
const MPV_REQUEST_ID_SUB_FONT = 111;
const MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW = 112;
const MPV_REQUEST_ID_OSD_HEIGHT = 113;
const MPV_REQUEST_ID_OSD_DIMENSIONS = 114;
const MPV_REQUEST_ID_SUBTEXT_ASS = 115;
const MPV_REQUEST_ID_SUB_SPACING = 116;
const MPV_REQUEST_ID_SUB_BOLD = 117;
const MPV_REQUEST_ID_SUB_ITALIC = 118;
const MPV_REQUEST_ID_SUB_BORDER_SIZE = 119;
const MPV_REQUEST_ID_SUB_SHADOW_OFFSET = 120;
const MPV_REQUEST_ID_SUB_ASS_OVERRIDE = 121;
const MPV_REQUEST_ID_SUB_USE_MARGINS = 122;
const MPV_REQUEST_ID_TRACK_LIST_SECONDARY = 200;
const MPV_REQUEST_ID_TRACK_LIST_AUDIO = 201;

class MpvIpcClient implements MpvClient {
  private socketPath: string;
  public socket: net.Socket | null = null;
  private buffer = "";
  public connected = false;
  private connecting = false;
  private reconnectAttempt = 0;
  private firstConnection = true;
  private hasConnectedOnce = false;
  public currentVideoPath = "";
  public currentTimePos = 0;
  public currentSubStart = 0;
  public currentSubEnd = 0;
  public currentSubText = "";
  public currentSecondarySubText = "";
  public currentAudioStreamIndex: number | null = null;
  private currentAudioTrackId: number | null = null;
  private pauseAtTime: number | null = null;
  private pendingPauseAtSubEnd = false;
  private nextDynamicRequestId = 1000;
  private pendingRequests = new Map<number, (message: MpvMessage) => void>();

  constructor(socketPath: string) {
    this.socketPath = socketPath;
  }

  setSocketPath(socketPath: string): void {
    this.socketPath = socketPath;
  }

  connect(): void {
    if (this.connected || this.connecting) {
      return;
    }

    if (this.socket) {
      this.socket.destroy();
    }

    this.connecting = true;
    this.socket = new net.Socket();

    this.socket.on("connect", () => {
      console.log("Connected to MPV socket");
      this.connected = true;
      this.connecting = false;
      this.reconnectAttempt = 0;
      this.hasConnectedOnce = true;
      this.subscribeToProperties();
      this.getInitialState();

      const shouldAutoStart =
        autoStartOverlay || getResolvedConfig().auto_start_overlay === true;
      if (this.firstConnection && shouldAutoStart) {
        console.log("Auto-starting overlay, hiding mpv subtitles");
        setTimeout(() => {
          setOverlayVisible(true);
        }, 100);
      } else if (shouldBindVisibleOverlayToMpvSubVisibility()) {
        this.setSubVisibility(!visibleOverlayVisible);
      }

      this.firstConnection = false;
    });

    this.socket.on("data", (data: Buffer) => {
      this.buffer += data.toString();
      this.processBuffer();
    });

    this.socket.on("error", (err: Error) => {
      console.error("MPV socket error:", err.message);
      this.connected = false;
      this.connecting = false;
      this.failPendingRequests();
    });

    this.socket.on("close", () => {
      console.log("MPV socket closed");
      this.connected = false;
      this.connecting = false;
      this.failPendingRequests();
      this.scheduleReconnect();
    });

    this.socket.connect(this.socketPath);
  }

  private scheduleReconnect(): void {
    if (reconnectTimer) {
      clearTimeout(reconnectTimer);
    }
    const attempt = this.reconnectAttempt++;
    let delay: number;
    if (this.hasConnectedOnce) {
      if (attempt < 2) {
        delay = 1000;
      } else if (attempt < 4) {
        delay = 2000;
      } else if (attempt < 7) {
        delay = 5000;
      } else {
        delay = 10000;
      }
    } else {
      if (attempt < 2) {
        delay = 200;
      } else if (attempt < 4) {
        delay = 500;
      } else if (attempt < 6) {
        delay = 1000;
      } else {
        delay = 2000;
      }
    }
    reconnectTimer = setTimeout(() => {
      console.log(
        `Attempting to reconnect to MPV (attempt ${attempt + 1}, delay ${delay}ms)...`,
      );
      this.connect();
    }, delay);
  }

  private processBuffer(): void {
    const lines = this.buffer.split("\n");
    this.buffer = lines.pop() || "";

    for (const line of lines) {
      if (!line.trim()) continue;
      try {
        const msg = JSON.parse(line) as MpvMessage;
        this.handleMessage(msg);
      } catch (e) {
        console.error("Failed to parse MPV message:", line, e);
      }
    }
  }

  private async handleMessage(msg: MpvMessage): Promise<void> {
    if (msg.event === "property-change") {
      if (msg.name === "sub-text") {
        currentSubText = (msg.data as string) || "";
        this.currentSubText = currentSubText;
        if (
          subtitleTimingTracker &&
          this.currentSubStart !== undefined &&
          this.currentSubEnd !== undefined
        ) {
          subtitleTimingTracker.recordSubtitle(
            currentSubText,
            this.currentSubStart,
            this.currentSubEnd,
          );
        }
        subtitleWsService.broadcast(currentSubText);
        if (getOverlayWindows().length > 0) {
          const subtitleData = await tokenizeSubtitle(currentSubText);
          broadcastToOverlayWindows("subtitle:set", subtitleData);
        }
      } else if (msg.name === "sub-text-ass") {
        currentSubAssText = (msg.data as string) || "";
        broadcastToOverlayWindows("subtitle-ass:set", currentSubAssText);
      } else if (msg.name === "sub-start") {
        this.currentSubStart = (msg.data as number) || 0;
        if (subtitleTimingTracker && currentSubText) {
          subtitleTimingTracker.recordSubtitle(
            currentSubText,
            this.currentSubStart,
            this.currentSubEnd,
          );
        }
      } else if (msg.name === "sub-end") {
        this.currentSubEnd = (msg.data as number) || 0;
        if (this.pendingPauseAtSubEnd && this.currentSubEnd > 0) {
          this.pauseAtTime = this.currentSubEnd;
          this.pendingPauseAtSubEnd = false;
          this.send({ command: ["set_property", "pause", false] });
        }
        if (subtitleTimingTracker && currentSubText) {
          subtitleTimingTracker.recordSubtitle(
            currentSubText,
            this.currentSubStart,
            this.currentSubEnd,
          );
        }
      } else if (msg.name === "secondary-sub-text") {
        this.currentSecondarySubText = (msg.data as string) || "";
        broadcastToOverlayWindows(
          "secondary-subtitle:set",
          this.currentSecondarySubText,
        );
      } else if (msg.name === "aid") {
        this.currentAudioTrackId =
          typeof msg.data === "number" ? (msg.data as number) : null;
        this.syncCurrentAudioStreamIndex();
      } else if (msg.name === "time-pos") {
        this.currentTimePos = (msg.data as number) || 0;
        if (
          this.pauseAtTime !== null &&
          this.currentTimePos >= this.pauseAtTime
        ) {
          this.pauseAtTime = null;
          this.send({ command: ["set_property", "pause", true] });
        }
      } else if (msg.name === "path") {
        this.currentVideoPath = (msg.data as string) || "";
        updateCurrentMediaPath(msg.data);
        this.autoLoadSecondarySubTrack();
        this.syncCurrentAudioStreamIndex();
      } else if (msg.name === "sub-pos") {
        updateMpvSubtitleRenderMetrics({ subPos: msg.data as number });
      } else if (msg.name === "sub-font-size") {
        updateMpvSubtitleRenderMetrics({ subFontSize: msg.data as number });
      } else if (msg.name === "sub-scale") {
        updateMpvSubtitleRenderMetrics({ subScale: msg.data as number });
      } else if (msg.name === "sub-margin-y") {
        updateMpvSubtitleRenderMetrics({ subMarginY: msg.data as number });
      } else if (msg.name === "sub-margin-x") {
        updateMpvSubtitleRenderMetrics({ subMarginX: msg.data as number });
      } else if (msg.name === "sub-font") {
        updateMpvSubtitleRenderMetrics({ subFont: msg.data as string });
      } else if (msg.name === "sub-spacing") {
        updateMpvSubtitleRenderMetrics({ subSpacing: msg.data as number });
      } else if (msg.name === "sub-bold") {
        updateMpvSubtitleRenderMetrics({
          subBold: asBoolean(msg.data, mpvSubtitleRenderMetrics.subBold),
        });
      } else if (msg.name === "sub-italic") {
        updateMpvSubtitleRenderMetrics({
          subItalic: asBoolean(msg.data, mpvSubtitleRenderMetrics.subItalic),
        });
      } else if (msg.name === "sub-border-size") {
        updateMpvSubtitleRenderMetrics({
          subBorderSize: msg.data as number,
        });
      } else if (msg.name === "sub-shadow-offset") {
        updateMpvSubtitleRenderMetrics({
          subShadowOffset: msg.data as number,
        });
      } else if (msg.name === "sub-ass-override") {
        updateMpvSubtitleRenderMetrics({
          subAssOverride: msg.data as string,
        });
      } else if (msg.name === "sub-scale-by-window") {
        updateMpvSubtitleRenderMetrics({
          subScaleByWindow: asBoolean(
            msg.data,
            mpvSubtitleRenderMetrics.subScaleByWindow,
          ),
        });
      } else if (msg.name === "sub-use-margins") {
        updateMpvSubtitleRenderMetrics({
          subUseMargins: asBoolean(
            msg.data,
            mpvSubtitleRenderMetrics.subUseMargins,
          ),
        });
      } else if (msg.name === "osd-height") {
        updateMpvSubtitleRenderMetrics({ osdHeight: msg.data as number });
      } else if (msg.name === "osd-dimensions") {
        const dims = msg.data as Record<string, unknown> | null;
        if (!dims) {
          updateMpvSubtitleRenderMetrics({ osdDimensions: null });
        } else {
          updateMpvSubtitleRenderMetrics({
            osdDimensions: {
              w: asFiniteNumber(dims.w, 0),
              h: asFiniteNumber(dims.h, 0),
              ml: asFiniteNumber(dims.ml, 0),
              mr: asFiniteNumber(dims.mr, 0),
              mt: asFiniteNumber(dims.mt, 0),
              mb: asFiniteNumber(dims.mb, 0),
            },
          });
        }
      }
    } else if (msg.request_id) {
      const pending = this.pendingRequests.get(msg.request_id);
      if (pending) {
        this.pendingRequests.delete(msg.request_id);
        pending(msg);
        return;
      }

      if (msg.data === undefined) {
        return;
      }

      if (msg.request_id === MPV_REQUEST_ID_TRACK_LIST_SECONDARY) {
        const tracks = msg.data as Array<{
          type: string;
          lang?: string;
          id: number;
        }>;
        if (Array.isArray(tracks)) {
          const config = getResolvedConfig();
          const languages = config.secondarySub?.secondarySubLanguages || [];
          const subTracks = tracks.filter((t) => t.type === "sub");
          for (const lang of languages) {
            const match = subTracks.find((t) => t.lang === lang);
            if (match) {
              this.send({
                command: ["set_property", "secondary-sid", match.id],
              });
              showMpvOsd(`Secondary subtitle: ${lang} (track ${match.id})`);
              break;
            }
          }
        }
      } else if (msg.request_id === MPV_REQUEST_ID_TRACK_LIST_AUDIO) {
        this.updateCurrentAudioStreamIndex(
          msg.data as Array<{
            type?: string;
            id?: number;
            selected?: boolean;
            "ff-index"?: number;
          }>,
        );
      } else if (msg.request_id === MPV_REQUEST_ID_SUBTEXT) {
        currentSubText = (msg.data as string) || "";
        if (mpvClient) {
          mpvClient.currentSubText = currentSubText;
        }
        subtitleWsService.broadcast(currentSubText);
        if (getOverlayWindows().length > 0) {
          tokenizeSubtitle(currentSubText).then((subtitleData) => {
            broadcastToOverlayWindows("subtitle:set", subtitleData);
          });
        }
      } else if (msg.request_id === MPV_REQUEST_ID_SUBTEXT_ASS) {
        currentSubAssText = (msg.data as string) || "";
        broadcastToOverlayWindows("subtitle-ass:set", currentSubAssText);
      } else if (msg.request_id === MPV_REQUEST_ID_PATH) {
        updateCurrentMediaPath(msg.data);
      } else if (msg.request_id === MPV_REQUEST_ID_AID) {
        this.currentAudioTrackId =
          typeof msg.data === "number" ? (msg.data as number) : null;
        this.syncCurrentAudioStreamIndex();
      } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUBTEXT) {
        this.currentSecondarySubText = (msg.data as string) || "";
        broadcastToOverlayWindows(
          "secondary-subtitle:set",
          this.currentSecondarySubText,
        );
      } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY) {
        if (!shouldBindVisibleOverlayToMpvSubVisibility()) {
          previousSecondarySubVisibility = null;
          return;
        }
        previousSecondarySubVisibility =
          msg.data === true || msg.data === "yes";
        this.send({
          command: ["set_property", "secondary-sub-visibility", "no"],
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_POS) {
        updateMpvSubtitleRenderMetrics({ subPos: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_FONT_SIZE) {
        updateMpvSubtitleRenderMetrics({ subFontSize: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SCALE) {
        updateMpvSubtitleRenderMetrics({ subScale: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_MARGIN_Y) {
        updateMpvSubtitleRenderMetrics({ subMarginY: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_MARGIN_X) {
        updateMpvSubtitleRenderMetrics({ subMarginX: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_FONT) {
        updateMpvSubtitleRenderMetrics({ subFont: msg.data as string });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SPACING) {
        updateMpvSubtitleRenderMetrics({ subSpacing: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_BOLD) {
        updateMpvSubtitleRenderMetrics({
          subBold: asBoolean(msg.data, mpvSubtitleRenderMetrics.subBold),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_ITALIC) {
        updateMpvSubtitleRenderMetrics({
          subItalic: asBoolean(msg.data, mpvSubtitleRenderMetrics.subItalic),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_BORDER_SIZE) {
        updateMpvSubtitleRenderMetrics({
          subBorderSize: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SHADOW_OFFSET) {
        updateMpvSubtitleRenderMetrics({
          subShadowOffset: msg.data as number,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_ASS_OVERRIDE) {
        updateMpvSubtitleRenderMetrics({
          subAssOverride: msg.data as string,
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW) {
        updateMpvSubtitleRenderMetrics({
          subScaleByWindow: asBoolean(
            msg.data,
            mpvSubtitleRenderMetrics.subScaleByWindow,
          ),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_SUB_USE_MARGINS) {
        updateMpvSubtitleRenderMetrics({
          subUseMargins: asBoolean(
            msg.data,
            mpvSubtitleRenderMetrics.subUseMargins,
          ),
        });
      } else if (msg.request_id === MPV_REQUEST_ID_OSD_HEIGHT) {
        updateMpvSubtitleRenderMetrics({ osdHeight: msg.data as number });
      } else if (msg.request_id === MPV_REQUEST_ID_OSD_DIMENSIONS) {
        const dims = msg.data as Record<string, unknown> | null;
        if (!dims) {
          updateMpvSubtitleRenderMetrics({ osdDimensions: null });
        } else {
          updateMpvSubtitleRenderMetrics({
            osdDimensions: {
              w: asFiniteNumber(dims.w, 0),
              h: asFiniteNumber(dims.h, 0),
              ml: asFiniteNumber(dims.ml, 0),
              mr: asFiniteNumber(dims.mr, 0),
              mt: asFiniteNumber(dims.mt, 0),
              mb: asFiniteNumber(dims.mb, 0),
            },
          });
        }
      }
    }
  }

  private autoLoadSecondarySubTrack(): void {
    const config = getResolvedConfig();
    if (!config.secondarySub?.autoLoadSecondarySub) return;
    const languages = config.secondarySub.secondarySubLanguages;
    if (!languages || languages.length === 0) return;

    setTimeout(() => {
      this.send({
        command: ["get_property", "track-list"],
        request_id: MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
      });
    }, 500);
  }

  private syncCurrentAudioStreamIndex(): void {
    this.send({
      command: ["get_property", "track-list"],
      request_id: MPV_REQUEST_ID_TRACK_LIST_AUDIO,
    });
  }

  private updateCurrentAudioStreamIndex(
    tracks: Array<{
      type?: string;
      id?: number;
      selected?: boolean;
      "ff-index"?: number;
    }>,
  ): void {
    if (!Array.isArray(tracks)) {
      this.currentAudioStreamIndex = null;
      return;
    }

    const audioTracks = tracks.filter((track) => track.type === "audio");
    const activeTrack =
      audioTracks.find((track) => track.id === this.currentAudioTrackId) ||
      audioTracks.find((track) => track.selected === true);

    const ffIndex = activeTrack?.["ff-index"];
    this.currentAudioStreamIndex =
      typeof ffIndex === "number" && Number.isInteger(ffIndex) && ffIndex >= 0
        ? ffIndex
        : null;
  }

  send(command: { command: unknown[]; request_id?: number }): boolean {
    if (!this.connected || !this.socket) {
      return false;
    }
    const msg = JSON.stringify(command) + "\n";
    this.socket.write(msg);
    return true;
  }

  request(command: unknown[]): Promise<MpvMessage> {
    return new Promise((resolve, reject) => {
      if (!this.connected || !this.socket) {
        reject(new Error("MPV not connected"));
        return;
      }

      const requestId = this.nextDynamicRequestId++;
      this.pendingRequests.set(requestId, resolve);
      const sent = this.send({ command, request_id: requestId });
      if (!sent) {
        this.pendingRequests.delete(requestId);
        reject(new Error("Failed to send MPV request"));
        return;
      }

      setTimeout(() => {
        if (this.pendingRequests.delete(requestId)) {
          reject(new Error("MPV request timed out"));
        }
      }, 4000);
    });
  }

  async requestProperty(name: string): Promise<unknown> {
    const response = await this.request(["get_property", name]);
    if (response.error && response.error !== "success") {
      throw new Error(
        `Failed to read MPV property '${name}': ${response.error}`,
      );
    }
    return response.data;
  }

  private failPendingRequests(): void {
    for (const [requestId, resolve] of this.pendingRequests.entries()) {
      resolve({ request_id: requestId, error: "disconnected" });
    }
    this.pendingRequests.clear();
  }

  private subscribeToProperties(): void {
    this.send({ command: ["observe_property", 1, "sub-text"] });
    this.send({ command: ["observe_property", 2, "path"] });
    this.send({ command: ["observe_property", 3, "sub-start"] });
    this.send({ command: ["observe_property", 4, "sub-end"] });
    this.send({ command: ["observe_property", 5, "time-pos"] });
    this.send({ command: ["observe_property", 6, "secondary-sub-text"] });
    this.send({ command: ["observe_property", 7, "aid"] });
    this.send({ command: ["observe_property", 8, "sub-pos"] });
    this.send({ command: ["observe_property", 9, "sub-font-size"] });
    this.send({ command: ["observe_property", 10, "sub-scale"] });
    this.send({ command: ["observe_property", 11, "sub-margin-y"] });
    this.send({ command: ["observe_property", 12, "sub-margin-x"] });
    this.send({ command: ["observe_property", 13, "sub-font"] });
    this.send({ command: ["observe_property", 14, "sub-spacing"] });
    this.send({ command: ["observe_property", 15, "sub-bold"] });
    this.send({ command: ["observe_property", 16, "sub-italic"] });
    this.send({ command: ["observe_property", 17, "sub-scale-by-window"] });
    this.send({ command: ["observe_property", 18, "osd-height"] });
    this.send({ command: ["observe_property", 19, "osd-dimensions"] });
    this.send({ command: ["observe_property", 20, "sub-text-ass"] });
    this.send({ command: ["observe_property", 21, "sub-border-size"] });
    this.send({ command: ["observe_property", 22, "sub-shadow-offset"] });
    this.send({ command: ["observe_property", 23, "sub-ass-override"] });
    this.send({ command: ["observe_property", 24, "sub-use-margins"] });
  }

  private getInitialState(): void {
    this.send({
      command: ["get_property", "sub-text"],
      request_id: MPV_REQUEST_ID_SUBTEXT,
    });
    this.send({
      command: ["get_property", "sub-text-ass"],
      request_id: MPV_REQUEST_ID_SUBTEXT_ASS,
    });
    this.send({
      command: ["get_property", "path"],
      request_id: MPV_REQUEST_ID_PATH,
    });
    this.send({
      command: ["get_property", "secondary-sub-text"],
      request_id: MPV_REQUEST_ID_SECONDARY_SUBTEXT,
    });
    this.send({
      command: ["get_property", "aid"],
      request_id: MPV_REQUEST_ID_AID,
    });
    this.send({
      command: ["get_property", "sub-pos"],
      request_id: MPV_REQUEST_ID_SUB_POS,
    });
    this.send({
      command: ["get_property", "sub-font-size"],
      request_id: MPV_REQUEST_ID_SUB_FONT_SIZE,
    });
    this.send({
      command: ["get_property", "sub-scale"],
      request_id: MPV_REQUEST_ID_SUB_SCALE,
    });
    this.send({
      command: ["get_property", "sub-margin-y"],
      request_id: MPV_REQUEST_ID_SUB_MARGIN_Y,
    });
    this.send({
      command: ["get_property", "sub-margin-x"],
      request_id: MPV_REQUEST_ID_SUB_MARGIN_X,
    });
    this.send({
      command: ["get_property", "sub-font"],
      request_id: MPV_REQUEST_ID_SUB_FONT,
    });
    this.send({
      command: ["get_property", "sub-spacing"],
      request_id: MPV_REQUEST_ID_SUB_SPACING,
    });
    this.send({
      command: ["get_property", "sub-bold"],
      request_id: MPV_REQUEST_ID_SUB_BOLD,
    });
    this.send({
      command: ["get_property", "sub-italic"],
      request_id: MPV_REQUEST_ID_SUB_ITALIC,
    });
    this.send({
      command: ["get_property", "sub-scale-by-window"],
      request_id: MPV_REQUEST_ID_SUB_SCALE_BY_WINDOW,
    });
    this.send({
      command: ["get_property", "osd-height"],
      request_id: MPV_REQUEST_ID_OSD_HEIGHT,
    });
    this.send({
      command: ["get_property", "osd-dimensions"],
      request_id: MPV_REQUEST_ID_OSD_DIMENSIONS,
    });
    this.send({
      command: ["get_property", "sub-border-size"],
      request_id: MPV_REQUEST_ID_SUB_BORDER_SIZE,
    });
    this.send({
      command: ["get_property", "sub-shadow-offset"],
      request_id: MPV_REQUEST_ID_SUB_SHADOW_OFFSET,
    });
    this.send({
      command: ["get_property", "sub-ass-override"],
      request_id: MPV_REQUEST_ID_SUB_ASS_OVERRIDE,
    });
    this.send({
      command: ["get_property", "sub-use-margins"],
      request_id: MPV_REQUEST_ID_SUB_USE_MARGINS,
    });
  }

  setSubVisibility(visible: boolean): void {
    this.send({
      command: ["set_property", "sub-visibility", visible ? "yes" : "no"],
    });
  }

  replayCurrentSubtitle(): void {
    this.pendingPauseAtSubEnd = true;
    this.send({ command: ["sub-seek", 0] });
  }

  playNextSubtitle(): void {
    this.pendingPauseAtSubEnd = true;
    this.send({ command: ["sub-seek", 1] });
  }
}

async function tokenizeSubtitle(text: string): Promise<SubtitleData> {
  const displayText = text
    .replace(/\r\n/g, "\n")
    .replace(/\\N/g, "\n")
    .replace(/\\n/g, "\n")
    .trim();

  if (!displayText) {
    return { text, tokens: null };
  }

  const tokenizeText = displayText
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const yomitanTokens = await parseWithYomitanInternalParser(tokenizeText);
  if (yomitanTokens && yomitanTokens.length > 0) {
    return { text: displayText, tokens: yomitanTokens };
  }

  if (!mecabTokenizer) {
    return { text: displayText, tokens: null };
  }

  try {
    const rawTokens = await mecabTokenizer.tokenize(tokenizeText);

    if (rawTokens && rawTokens.length > 0) {
      const mergedTokens = mergeTokens(rawTokens);
      return { text: displayText, tokens: mergedTokens };
    }
  } catch (err) {
    console.error("Tokenization error:", (err as Error).message);
  }

  return { text: displayText, tokens: null };
}

interface YomitanParseHeadword {
  term?: unknown;
}

interface YomitanParseSegment {
  text?: unknown;
  reading?: unknown;
  headwords?: unknown;
}

interface YomitanParseResultItem {
  source?: unknown;
  index?: unknown;
  content?: unknown;
}

function extractYomitanHeadword(segment: YomitanParseSegment): string {
  const headwords = segment.headwords;
  if (!Array.isArray(headwords) || headwords.length === 0) {
    return "";
  }

  const firstGroup = headwords[0];
  if (!Array.isArray(firstGroup) || firstGroup.length === 0) {
    return "";
  }

  const firstHeadword = firstGroup[0] as YomitanParseHeadword;
  return typeof firstHeadword?.term === "string" ? firstHeadword.term : "";
}

function mapYomitanParseResultsToMergedTokens(
  parseResults: unknown,
): MergedToken[] | null {
  if (!Array.isArray(parseResults) || parseResults.length === 0) {
    return null;
  }

  const scanningItems = parseResults.filter((item) => {
    const resultItem = item as YomitanParseResultItem;
    return (
      resultItem &&
      resultItem.source === "scanning-parser" &&
      Array.isArray(resultItem.content)
    );
  }) as YomitanParseResultItem[];

  if (scanningItems.length === 0) {
    return null;
  }

  const primaryItem =
    scanningItems.find((item) => item.index === 0) || scanningItems[0];
  const content = primaryItem.content;
  if (!Array.isArray(content)) {
    return null;
  }

  const tokens: MergedToken[] = [];
  let charOffset = 0;

  for (const line of content) {
    if (!Array.isArray(line)) {
      continue;
    }

    let surface = "";
    let reading = "";
    let headword = "";

    for (const rawSegment of line) {
      const segment = rawSegment as YomitanParseSegment;
      if (!segment || typeof segment !== "object") {
        continue;
      }

      const segmentText = segment.text;
      if (typeof segmentText !== "string" || segmentText.length === 0) {
        continue;
      }

      surface += segmentText;

      if (typeof segment.reading === "string") {
        reading += segment.reading;
      }

      if (!headword) {
        headword = extractYomitanHeadword(segment);
      }
    }

    if (!surface) {
      continue;
    }

    const start = charOffset;
    const end = start + surface.length;
    charOffset = end;

    tokens.push({
      surface,
      reading,
      headword: headword || surface,
      startPos: start,
      endPos: end,
      partOfSpeech: PartOfSpeech.other,
      isMerged: true,
    });
  }

  return tokens.length > 0 ? tokens : null;
}

async function ensureYomitanParserWindow(): Promise<boolean> {
  if (!yomitanExt) {
    return false;
  }

  if (yomitanParserWindow && !yomitanParserWindow.isDestroyed()) {
    return true;
  }

  if (yomitanParserInitPromise) {
    return yomitanParserInitPromise;
  }

  yomitanParserInitPromise = (async () => {
    const parserWindow = new BrowserWindow({
      show: false,
      width: 800,
      height: 600,
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false,
        session: session.defaultSession,
      },
    });
    yomitanParserWindow = parserWindow;

    yomitanParserReadyPromise = new Promise((resolve, reject) => {
      parserWindow.webContents.once("did-finish-load", () => resolve());
      parserWindow.webContents.once(
        "did-fail-load",
        (_event, _errorCode, errorDescription) => {
          reject(new Error(errorDescription));
        },
      );
    });

    parserWindow.on("closed", () => {
      if (yomitanParserWindow === parserWindow) {
        yomitanParserWindow = null;
        yomitanParserReadyPromise = null;
      }
    });

    try {
      await parserWindow.loadURL(`chrome-extension://${yomitanExt.id}/search.html`);
      if (yomitanParserReadyPromise) {
        await yomitanParserReadyPromise;
      }
      return true;
    } catch (err) {
      console.error(
        "Failed to initialize Yomitan parser window:",
        (err as Error).message,
      );
      if (!parserWindow.isDestroyed()) {
        parserWindow.destroy();
      }
      if (yomitanParserWindow === parserWindow) {
        yomitanParserWindow = null;
        yomitanParserReadyPromise = null;
      }
      return false;
    } finally {
      yomitanParserInitPromise = null;
    }
  })();

  return yomitanParserInitPromise;
}

async function parseWithYomitanInternalParser(
  text: string,
): Promise<MergedToken[] | null> {
  if (!text || !yomitanExt) {
    return null;
  }

  const isReady = await ensureYomitanParserWindow();
  if (!isReady || !yomitanParserWindow || yomitanParserWindow.isDestroyed()) {
    return null;
  }

  const script = `
    (async () => {
      const invoke = (action, params) =>
        new Promise((resolve, reject) => {
          chrome.runtime.sendMessage({ action, params }, (response) => {
            if (chrome.runtime.lastError) {
              reject(new Error(chrome.runtime.lastError.message));
              return;
            }
            if (!response || typeof response !== "object") {
              reject(new Error("Invalid response from Yomitan backend"));
              return;
            }
            if (response.error) {
              reject(new Error(response.error.message || "Yomitan backend error"));
              return;
            }
            resolve(response.result);
          });
        });

      const optionsFull = await invoke("optionsGetFull", undefined);
      const profileIndex = optionsFull.profileCurrent;
      const scanLength =
        optionsFull.profiles?.[profileIndex]?.options?.scanning?.length ?? 40;

      return await invoke("parseText", {
        text: ${JSON.stringify(text)},
        optionsContext: { index: profileIndex },
        scanLength,
        useInternalParser: true,
        useMecabParser: false
      });
    })();
  `;

  try {
    const parseResults = await yomitanParserWindow.webContents.executeJavaScript(
      script,
      true,
    );
    return mapYomitanParseResultsToMergedTokens(parseResults);
  } catch (err) {
    console.error("Yomitan parser request failed:", (err as Error).message);
    return null;
  }
}

function updateOverlayBounds(geometry: WindowGeometry): void {
  if (!geometry) return;
  for (const window of getOverlayWindows()) {
    window.setBounds({
      x: geometry.x,
      y: geometry.y,
      width: geometry.width,
      height: geometry.height,
    });
  }
}

function ensureOverlayWindowLevel(window: BrowserWindow): void {
  if (process.platform === "darwin") {
    window.setAlwaysOnTop(true, "screen-saver", 1);
    window.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
    window.setFullScreenable(false);
    return;
  }
  window.setAlwaysOnTop(true);
}

function enforceOverlayLayerOrder(): void {
  if (!visibleOverlayVisible || !invisibleOverlayVisible) return;
  if (!mainWindow || mainWindow.isDestroyed()) return;
  if (!invisibleWindow || invisibleWindow.isDestroyed()) return;

  ensureOverlayWindowLevel(mainWindow);
  mainWindow.moveTop();
}

function ensureExtensionCopy(sourceDir: string): string {
  if (process.platform === "win32") {
    return sourceDir;
  }

  const extensionsRoot = path.join(USER_DATA_PATH, "extensions");
  const targetDir = path.join(extensionsRoot, "yomitan");

  const sourceManifest = path.join(sourceDir, "manifest.json");
  const targetManifest = path.join(targetDir, "manifest.json");

  let shouldCopy = !fs.existsSync(targetDir);
  if (
    !shouldCopy &&
    fs.existsSync(sourceManifest) &&
    fs.existsSync(targetManifest)
  ) {
    try {
      const sourceVersion = (
        JSON.parse(fs.readFileSync(sourceManifest, "utf-8")) as {
          version: string;
        }
      ).version;
      const targetVersion = (
        JSON.parse(fs.readFileSync(targetManifest, "utf-8")) as {
          version: string;
        }
      ).version;
      shouldCopy = sourceVersion !== targetVersion;
    } catch (e) {
      shouldCopy = true;
    }
  }

  if (shouldCopy) {
    fs.mkdirSync(extensionsRoot, { recursive: true });
    fs.rmSync(targetDir, { recursive: true, force: true });
    fs.cpSync(sourceDir, targetDir, { recursive: true });
    console.log(`Copied yomitan extension to ${targetDir}`);
  }

  return targetDir;
}

async function loadYomitanExtension(): Promise<Extension | null> {
  const searchPaths = [
    path.join(__dirname, "..", "vendor", "yomitan"),
    path.join(process.resourcesPath, "yomitan"),
    "/usr/share/SubMiner/yomitan",
    path.join(USER_DATA_PATH, "yomitan"),
  ];

  let extPath: string | null = null;
  for (const p of searchPaths) {
    if (fs.existsSync(p)) {
      extPath = p;
      break;
    }
  }


  if (!extPath) {
    console.error("Yomitan extension not found in any search path");
    console.error("Install Yomitan to one of:", searchPaths);
    return null;
  }

  extPath = ensureExtensionCopy(extPath);

  if (yomitanParserWindow && !yomitanParserWindow.isDestroyed()) {
    yomitanParserWindow.destroy();
  }
  yomitanParserWindow = null;
  yomitanParserReadyPromise = null;
  yomitanParserInitPromise = null;

  try {
    const extensions = session.defaultSession.extensions;
    if (extensions) {
      yomitanExt = await extensions.loadExtension(extPath, {
        allowFileAccess: true,
      });
    } else {
      yomitanExt = await session.defaultSession.loadExtension(extPath, {
        allowFileAccess: true,
      });
    }
    return yomitanExt;
  } catch (err) {
    console.error("Failed to load Yomitan extension:", (err as Error).message);
    console.error("Full error:", err);
    return null;
  }
}

function createOverlayWindow(kind: "visible" | "invisible"): BrowserWindow {
  const window = new BrowserWindow({
    show: false,
    width: 800,
    height: 600,
    x: 0,
    y: 0,
    transparent: true,
    frame: false,
    alwaysOnTop: true,
    skipTaskbar: true,
    resizable: false,
    hasShadow: false,
    focusable: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: true,
      additionalArguments: [`--overlay-layer=${kind}`],
    },
  });

  ensureOverlayWindowLevel(window);

  const htmlPath = path.join(__dirname, "renderer", "index.html");

  window
    .loadFile(htmlPath, {
      query: { layer: kind === "visible" ? "visible" : "invisible" },
    })
    .catch((err) => {
      console.error("Failed to load HTML file:", err);
    });

  window.webContents.on(
    "did-fail-load",
    (_event, errorCode, errorDescription, validatedURL) => {
      console.error(
        "Page failed to load:",
        errorCode,
        errorDescription,
        validatedURL,
      );
    },
  );

  window.webContents.on("did-finish-load", () => {
    broadcastRuntimeOptionsChanged();
    window.webContents.send(
      "overlay-debug-visualization:set",
      overlayDebugVisualizationEnabled,
    );
  });

  if (kind === "visible") {
    window.webContents.on("devtools-opened", () => {
      setOverlayDebugVisualizationEnabled(true);
    });
    window.webContents.on("devtools-closed", () => {
      setOverlayDebugVisualizationEnabled(false);
    });
  }

  window.webContents.on("before-input-event", (event, input) => {
    const isOverlayVisible =
      kind === "visible" ? visibleOverlayVisible : invisibleOverlayVisible;
    if (!isOverlayVisible) return;
    if (!tryHandleOverlayShortcutLocalFallback(input)) return;
    event.preventDefault();
  });

  window.hide();

  window.on("closed", () => {
    if (kind === "visible") {
      mainWindow = null;
    } else {
      invisibleWindow = null;
    }
  });

  window.on("blur", () => {
    if (!window.isDestroyed()) {
      ensureOverlayWindowLevel(window);
    }
  });

  if (isDev && kind === "visible") {
    window.webContents.openDevTools({ mode: "detach" });
  }

  return window;
}

function createMainWindow(): BrowserWindow {
  mainWindow = createOverlayWindow("visible");
  return mainWindow;
}

function createInvisibleWindow(): BrowserWindow {
  invisibleWindow = createOverlayWindow("invisible");
  return invisibleWindow;
}

function initializeOverlayRuntime(): void {
  if (overlayRuntimeInitialized) {
    return;
  }

  createMainWindow();
  createInvisibleWindow();
  invisibleOverlayVisible = getInitialInvisibleOverlayVisibility();
  registerGlobalShortcuts();

  windowTracker = createWindowTracker(backendOverride);
  if (windowTracker) {
    windowTracker.onGeometryChange = (geometry: WindowGeometry) => {
      updateOverlayBounds(geometry);
    };
    windowTracker.onWindowFound = (geometry: WindowGeometry) => {
      updateOverlayBounds(geometry);
      if (visibleOverlayVisible) {
        updateVisibleOverlayVisibility();
      }
      if (invisibleOverlayVisible) {
        updateInvisibleOverlayVisibility();
      }
    };
    windowTracker.onWindowLost = () => {
      for (const window of getOverlayWindows()) {
        window.hide();
      }
      syncOverlayShortcuts();
    };
    windowTracker.start();
  }

  const config = getResolvedConfig();
  if (
    config.ankiConnect?.enabled &&
    subtitleTimingTracker &&
    mpvClient &&
    runtimeOptionsManager
  ) {
    const effectiveAnkiConfig =
      runtimeOptionsManager.getEffectiveAnkiConnectConfig(config.ankiConnect);
    ankiIntegration = new AnkiIntegration(
      effectiveAnkiConfig,
      subtitleTimingTracker,
      mpvClient,
      (text: string) => {
        if (mpvClient) {
          mpvClient.send({
            command: ["show-text", text, "3000"],
          });
        }
      },
      showDesktopNotification,
      createFieldGroupingCallback(),
    );
    ankiIntegration.start();
  }

  overlayRuntimeInitialized = true;
  updateVisibleOverlayVisibility();
  updateInvisibleOverlayVisibility();
}

function openYomitanSettings(): void {
  openYomitanSettingsWindow({
    yomitanExt,
    getExistingWindow: () => yomitanSettingsWindow,
    setWindow: (window) => (yomitanSettingsWindow = window),
  });
}

function registerGlobalShortcuts(): void {
  registerGlobalShortcutsService({
    shortcuts: getConfiguredShortcuts(),
    onToggleVisibleOverlay: () => toggleVisibleOverlay(),
    onToggleInvisibleOverlay: () => toggleInvisibleOverlay(),
    onOpenYomitanSettings: () => openYomitanSettings(),
    isDev,
    getMainWindow: () => mainWindow,
  });
}

function getConfiguredShortcuts() { return resolveConfiguredShortcuts(getResolvedConfig(), DEFAULT_CONFIG); }

function tryHandleOverlayShortcutLocalFallback(input: Electron.Input): boolean {
  const shortcuts = getConfiguredShortcuts();
  return runOverlayShortcutLocalFallback(
    input,
    shortcuts,
    shortcutMatchesInputForLocalFallback,
    {
      openRuntimeOptions: () => {
        openRuntimeOptionsPalette();
      },
      openJimaku: () => {
        sendToVisibleOverlay("jimaku:open");
      },
      markAudioCard: () => {
        markLastCardAsAudioCard().catch((err) => {
          console.error("markLastCardAsAudioCard failed:", err);
          showMpvOsd(`Audio card failed: ${(err as Error).message}`);
        });
      },
      copySubtitleMultiple: (timeoutMs) => {
        startPendingMultiCopy(timeoutMs);
      },
      copySubtitle: () => {
        copyCurrentSubtitle();
      },
      toggleSecondarySub: () => cycleSecondarySubMode(),
      updateLastCardFromClipboard: () => {
        updateLastCardFromClipboard().catch((err) => {
          console.error("updateLastCardFromClipboard failed:", err);
          showMpvOsd(`Update failed: ${(err as Error).message}`);
        });
      },
      triggerFieldGrouping: () => {
        triggerFieldGrouping().catch((err) => {
          console.error("triggerFieldGrouping failed:", err);
          showMpvOsd(`Field grouping failed: ${(err as Error).message}`);
        });
      },
      triggerSubsync: () => {
        triggerSubsyncFromConfig().catch((err) => {
          console.error("triggerSubsyncFromConfig failed:", err);
          showMpvOsd(`Subsync failed: ${(err as Error).message}`);
        });
      },
      mineSentence: () => {
        mineSentenceCard().catch((err) => {
          console.error("mineSentenceCard failed:", err);
          showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
        });
      },
      mineSentenceMultiple: (timeoutMs) => {
        startPendingMineSentenceMultiple(timeoutMs);
      },
    },
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
  shortcutsRegistered = registerOverlayShortcutsService(shortcuts, {
    copySubtitle: () => {
      copyCurrentSubtitle();
    },
    copySubtitleMultiple: (timeoutMs) => {
      startPendingMultiCopy(timeoutMs);
    },
    updateLastCardFromClipboard: () => {
      updateLastCardFromClipboard().catch((err) => {
        console.error("updateLastCardFromClipboard failed:", err);
        showMpvOsd(`Update failed: ${(err as Error).message}`);
      });
    },
    triggerFieldGrouping: () => {
      triggerFieldGrouping().catch((err) => {
        console.error("triggerFieldGrouping failed:", err);
        showMpvOsd(`Field grouping failed: ${(err as Error).message}`);
      });
    },
    triggerSubsync: () => {
      triggerSubsyncFromConfig().catch((err) => {
        console.error("triggerSubsyncFromConfig failed:", err);
        showMpvOsd(`Subsync failed: ${(err as Error).message}`);
      });
    },
    mineSentence: () => {
      mineSentenceCard().catch((err) => {
        console.error("mineSentenceCard failed:", err);
        showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
      });
    },
    mineSentenceMultiple: (timeoutMs) => {
      startPendingMineSentenceMultiple(timeoutMs);
    },
    toggleSecondarySub: () => cycleSecondarySubMode(),
    markAudioCard: () => {
      markLastCardAsAudioCard().catch((err) => {
        console.error("markLastCardAsAudioCard failed:", err);
        showMpvOsd(`Audio card failed: ${(err as Error).message}`);
      });
    },
    openRuntimeOptions: () => {
      openRuntimeOptionsPalette();
    },
    openJimaku: () => {
      sendToVisibleOverlay("jimaku:open");
    },
  });
}

function unregisterOverlayShortcuts(): void {
  if (!shortcutsRegistered) return;

  cancelPendingMultiCopy();
  cancelPendingMineSentenceMultiple();

  unregisterOverlayShortcutsService(getConfiguredShortcuts());

  shortcutsRegistered = false;
}

function shouldOverlayShortcutsBeActive(): boolean {
  return overlayRuntimeInitialized;
}

function syncOverlayShortcuts(): void {
  if (shouldOverlayShortcutsBeActive()) {
    registerOverlayShortcuts();
  } else {
    unregisterOverlayShortcuts();
  }
}

function refreshOverlayShortcuts(): void {
  unregisterOverlayShortcuts();
  syncOverlayShortcuts();
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
  if (!invisibleWindow || invisibleWindow.isDestroyed()) return;
  if (visibleOverlayVisible) {
    invisibleWindow.setIgnoreMouseEvents(true, { forward: true });
  } else if (invisibleOverlayVisible) {
    invisibleWindow.setIgnoreMouseEvents(false);
  }
}

function setVisibleOverlayVisible(visible: boolean): void {
  visibleOverlayVisible = visible;
  updateVisibleOverlayVisibility();
  updateInvisibleOverlayVisibility();
  syncInvisibleOverlayMousePassthrough();
  if (
    shouldBindVisibleOverlayToMpvSubVisibility() &&
    mpvClient &&
    mpvClient.connected
  ) {
    mpvClient.setSubVisibility(!visible);
  }
}

function setInvisibleOverlayVisible(visible: boolean): void {
  invisibleOverlayVisible = visible;
  updateInvisibleOverlayVisibility();
  syncInvisibleOverlayMousePassthrough();
}

function toggleVisibleOverlay(): void {
  setVisibleOverlayVisible(!visibleOverlayVisible);
}

function toggleInvisibleOverlay(): void {
  setInvisibleOverlayVisible(!invisibleOverlayVisible);
}

function setOverlayVisible(visible: boolean): void {
  setVisibleOverlayVisible(visible);
}

function toggleOverlay(): void {
  toggleVisibleOverlay();
}

function handleOverlayModalClosed(modal: OverlayHostedModal): void {
  if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
  restoreVisibleOverlayOnModalClose.delete(modal);
  if (restoreVisibleOverlayOnModalClose.size === 0) {
    setVisibleOverlayVisible(false);
  }
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
  return async (
    data: KikuFieldGroupingRequestData,
  ): Promise<KikuFieldGroupingChoice> => {
    return new Promise((resolve) => {
      const previousVisibleOverlay = visibleOverlayVisible;
      const previousInvisibleOverlay = invisibleOverlayVisible;
      let settled = false;

      const finish = (choice: KikuFieldGroupingChoice): void => {
        if (settled) return;
        settled = true;
        fieldGroupingResolver = null;
        resolve(choice);

        if (!previousVisibleOverlay && visibleOverlayVisible) {
          setVisibleOverlayVisible(false);
        }
        if (invisibleOverlayVisible !== previousInvisibleOverlay) {
          setInvisibleOverlayVisible(previousInvisibleOverlay);
        }
      };

      fieldGroupingResolver = finish;
      if (!sendToVisibleOverlay("kiku:field-grouping-request", data)) {
        finish({
          keepNoteId: 0,
          deleteNoteId: 0,
          deleteDuplicate: true,
          cancelled: true,
        });
        return;
      }
      setTimeout(() => {
        if (!settled) {
          finish({
            keepNoteId: 0,
            deleteNoteId: 0,
            deleteDuplicate: true,
            cancelled: true,
          });
        }
      }, 90000);
    });
  };
}

function sendToVisibleOverlay(
  channel: string,
  payload?: unknown,
  options?: { restoreOnModalClose?: OverlayHostedModal },
): boolean {
  return sendToVisibleOverlayService({
    mainWindow,
    visibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    channel,
    payload,
    restoreOnModalClose: options?.restoreOnModalClose,
    addRestoreFlag: (modal) =>
      restoreVisibleOverlayOnModalClose.add(modal as OverlayHostedModal),
  });
}

registerAnkiJimakuIpcHandlers({
  setAnkiConnectEnabled: (enabled) => {
    configService.patchRawConfig({
      ankiConnect: {
        enabled,
      },
    });
    const config = getResolvedConfig();

    if (enabled && !ankiIntegration && subtitleTimingTracker && mpvClient) {
      const effectiveAnkiConfig = runtimeOptionsManager
        ? runtimeOptionsManager.getEffectiveAnkiConnectConfig(
            config.ankiConnect,
          )
        : config.ankiConnect;
      ankiIntegration = new AnkiIntegration(
        effectiveAnkiConfig,
        subtitleTimingTracker,
        mpvClient,
        (text: string) => {
          if (mpvClient) {
            mpvClient.send({
              command: ["show-text", text, "3000"],
            });
          }
        },
        showDesktopNotification,
        createFieldGroupingCallback(),
      );
      ankiIntegration.start();
      console.log("AnkiConnect integration enabled");
    } else if (!enabled && ankiIntegration) {
      ankiIntegration.destroy();
      ankiIntegration = null;
      console.log("AnkiConnect integration disabled");
    }

    broadcastRuntimeOptionsChanged();
  },
  clearAnkiHistory: () => {
    if (subtitleTimingTracker) {
      subtitleTimingTracker.cleanup();
      console.log("AnkiConnect subtitle timing history cleared");
    }
  },
  respondFieldGrouping: (choice) => {
    if (fieldGroupingResolver) {
      fieldGroupingResolver(choice);
      fieldGroupingResolver = null;
    }
  },
  buildKikuMergePreview: async (request) => {
    if (!ankiIntegration) {
      return { ok: false, error: "AnkiConnect integration not enabled" };
    }
    return ankiIntegration.buildFieldGroupingPreview(
      request.keepNoteId,
      request.deleteNoteId,
      request.deleteDuplicate,
    );
  },
  getJimakuMediaInfo: () => parseMediaInfo(currentMediaPath),
  searchJimakuEntries: async (query) => {
    console.log(`[jimaku] search-entries query: "${query.query}"`);
    const response = await jimakuFetchJson<JimakuEntry[]>(
      "/api/entries/search",
      {
        anime: true,
        query: query.query,
      },
    );
    if (!response.ok) return response;
    const maxResults = getJimakuMaxEntryResults();
    console.log(
      `[jimaku] search-entries returned ${response.data.length} results (capped to ${maxResults})`,
    );
    return { ok: true, data: response.data.slice(0, maxResults) };
  },
  listJimakuFiles: async (query) => {
    console.log(
      `[jimaku] list-files entryId=${query.entryId} episode=${query.episode ?? "all"}`,
    );
    const response = await jimakuFetchJson<JimakuFileEntry[]>(
      `/api/entries/${query.entryId}/files`,
      {
        episode: query.episode ?? undefined,
      },
    );
    if (!response.ok) return response;
    const sorted = sortJimakuFiles(
      response.data,
      getJimakuLanguagePreference(),
    );
    console.log(`[jimaku] list-files returned ${sorted.length} files`);
    return { ok: true, data: sorted };
  },
  resolveJimakuApiKey: () => resolveJimakuApiKey(),
  getCurrentMediaPath: () => currentMediaPath,
  isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath),
  downloadToFile: (url, destPath, headers) => downloadToFile(url, destPath, headers),
  onDownloadedSubtitle: (pathToSubtitle) => {
    if (mpvClient && mpvClient.connected) {
      mpvClient.send({ command: ["sub-add", pathToSubtitle, "select"] });
    }
  },
});
