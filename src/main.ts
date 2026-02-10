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
  ipcMain,
  globalShortcut,
  clipboard,
  shell,
  protocol,
  screen,
  IpcMainEvent,
  Extension,
  Notification,
  nativeImage,
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
import * as readline from "readline";
import * as childProcess from "child_process";
import WebSocket from "ws";
import { MecabTokenizer } from "./mecab-tokenizer";
import { mergeTokens } from "./token-merger";
import { createWindowTracker, BaseWindowTracker } from "./window-trackers";
import {
  Config,
  JimakuApiResponse,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuFilesQuery,
  JimakuMediaInfo,
  JimakuSearchQuery,
  JimakuDownloadQuery,
  JimakuConfig,
  JimakuLanguagePreference,
  SubtitleData,
  SubtitlePosition,
  Keybinding,
  WindowGeometry,
  SecondarySubMode,
  MpvClient,
  SubsyncConfig,
  SubsyncMode,
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
  ConfigService,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
  SPECIAL_COMMANDS,
} from "./config";

if (process.platform === "linux") {
  // Wayland requires the portal backend for reliable globalShortcut support.
  app.commandLine.appendSwitch("enable-features", "GlobalShortcutsPortal");
}

const DEFAULT_TEXTHOOKER_PORT = 5174;
let texthookerServer: http.Server | null = null;
let subtitleWebSocketServer: WebSocket.Server | null = null;
const CONFIG_DIR = path.join(os.homedir(), ".config", "SubMiner");
const USER_DATA_PATH = CONFIG_DIR;
const configService = new ConfigService(CONFIG_DIR);
const isDev =
  process.argv.includes("--dev") || process.argv.includes("--debug");

function getDefaultSocketPath(): string {
  if (process.platform === "win32") {
    return "\\\\.\\pipe\\subminer-socket";
  }
  return "/tmp/subminer-socket";
}

interface CliArgs {
  start: boolean;
  stop: boolean;
  toggle: boolean;
  toggleVisibleOverlay: boolean;
  toggleInvisibleOverlay: boolean;
  settings: boolean;
  show: boolean;
  hide: boolean;
  showVisibleOverlay: boolean;
  hideVisibleOverlay: boolean;
  showInvisibleOverlay: boolean;
  hideInvisibleOverlay: boolean;
  copySubtitle: boolean;
  copySubtitleMultiple: boolean;
  mineSentence: boolean;
  mineSentenceMultiple: boolean;
  updateLastCardFromClipboard: boolean;
  toggleSecondarySub: boolean;
  triggerFieldGrouping: boolean;
  triggerSubsync: boolean;
  markAudioCard: boolean;
  openRuntimeOptions: boolean;
  texthooker: boolean;
  help: boolean;
  autoStartOverlay: boolean;
  generateConfig: boolean;
  configPath?: string;
  backupOverwrite: boolean;
  socketPath?: string;
  backend?: string;
  texthookerPort?: number;
  verbose: boolean;
  logLevel?: "debug" | "info" | "warn" | "error";
}

type CliCommandSource = "initial" | "second-instance";

function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = {
    start: false,
    stop: false,
    toggle: false,
    toggleVisibleOverlay: false,
    toggleInvisibleOverlay: false,
    settings: false,
    show: false,
    hide: false,
    showVisibleOverlay: false,
    hideVisibleOverlay: false,
    showInvisibleOverlay: false,
    hideInvisibleOverlay: false,
    copySubtitle: false,
    copySubtitleMultiple: false,
    mineSentence: false,
    mineSentenceMultiple: false,
    updateLastCardFromClipboard: false,
    toggleSecondarySub: false,
    triggerFieldGrouping: false,
    triggerSubsync: false,
    markAudioCard: false,
    openRuntimeOptions: false,
    texthooker: false,
    help: false,
    autoStartOverlay: false,
    generateConfig: false,
    backupOverwrite: false,
    verbose: false,
  };

  const readValue = (value?: string): string | undefined => {
    if (!value) return undefined;
    if (value.startsWith("--")) return undefined;
    return value;
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith("--")) continue;

    if (arg === "--start") args.start = true;
    else if (arg === "--stop") args.stop = true;
    else if (arg === "--toggle") args.toggle = true;
    else if (arg === "--toggle-visible-overlay")
      args.toggleVisibleOverlay = true;
    else if (arg === "--toggle-invisible-overlay")
      args.toggleInvisibleOverlay = true;
    else if (arg === "--settings" || arg === "--yomitan") args.settings = true;
    else if (arg === "--show") args.show = true;
    else if (arg === "--hide") args.hide = true;
    else if (arg === "--show-visible-overlay") args.showVisibleOverlay = true;
    else if (arg === "--hide-visible-overlay") args.hideVisibleOverlay = true;
    else if (arg === "--show-invisible-overlay")
      args.showInvisibleOverlay = true;
    else if (arg === "--hide-invisible-overlay")
      args.hideInvisibleOverlay = true;
    else if (arg === "--copy-subtitle") args.copySubtitle = true;
    else if (arg === "--copy-subtitle-multiple") args.copySubtitleMultiple = true;
    else if (arg === "--mine-sentence") args.mineSentence = true;
    else if (arg === "--mine-sentence-multiple") args.mineSentenceMultiple = true;
    else if (arg === "--update-last-card-from-clipboard")
      args.updateLastCardFromClipboard = true;
    else if (arg === "--toggle-secondary-sub") args.toggleSecondarySub = true;
    else if (arg === "--trigger-field-grouping")
      args.triggerFieldGrouping = true;
    else if (arg === "--trigger-subsync") args.triggerSubsync = true;
    else if (arg === "--mark-audio-card") args.markAudioCard = true;
    else if (arg === "--open-runtime-options") args.openRuntimeOptions = true;
    else if (arg === "--texthooker") args.texthooker = true;
    else if (arg === "--auto-start-overlay") args.autoStartOverlay = true;
    else if (arg === "--generate-config") args.generateConfig = true;
    else if (arg === "--backup-overwrite") args.backupOverwrite = true;
    else if (arg === "--help") args.help = true;
    else if (arg === "--verbose") args.verbose = true;
    else if (arg.startsWith("--log-level=")) {
      const value = arg.split("=", 2)[1]?.toLowerCase();
      if (
        value === "debug" ||
        value === "info" ||
        value === "warn" ||
        value === "error"
      ) {
        args.logLevel = value;
      }
    } else if (arg === "--log-level") {
      const value = readValue(argv[i + 1])?.toLowerCase();
      if (
        value === "debug" ||
        value === "info" ||
        value === "warn" ||
        value === "error"
      ) {
        args.logLevel = value;
      }
    } else if (arg.startsWith("--config-path=")) {
      const value = arg.split("=", 2)[1];
      if (value) args.configPath = value;
    } else if (arg === "--config-path") {
      const value = readValue(argv[i + 1]);
      if (value) args.configPath = value;
    } else if (arg.startsWith("--socket=")) {
      const value = arg.split("=", 2)[1];
      if (value) args.socketPath = value;
    } else if (arg === "--socket") {
      const value = readValue(argv[i + 1]);
      if (value) args.socketPath = value;
    } else if (arg.startsWith("--backend=")) {
      const value = arg.split("=", 2)[1];
      if (value) args.backend = value;
    } else if (arg === "--backend") {
      const value = readValue(argv[i + 1]);
      if (value) args.backend = value;
    } else if (arg.startsWith("--port=")) {
      const value = Number(arg.split("=", 2)[1]);
      if (!Number.isNaN(value)) args.texthookerPort = value;
    } else if (arg === "--port") {
      const value = Number(readValue(argv[i + 1]));
      if (!Number.isNaN(value)) args.texthookerPort = value;
    }
  }

  return args;
}

function printHelp(): void {
  console.log(`
SubMiner CLI commands:
  --start               Start MPV IPC connection and overlay control loop
  --stop                Stop the running overlay app
  --toggle              Toggle visible subtitle overlay visibility (legacy alias)
  --toggle-visible-overlay    Toggle visible subtitle overlay visibility
  --toggle-invisible-overlay  Toggle invisible interactive overlay visibility
  --settings            Open Yomitan settings window
  --texthooker          Launch texthooker only (no overlay window)
  --show                Force show visible overlay (legacy alias)
  --hide                Force hide visible overlay (legacy alias)
  --show-visible-overlay       Force show visible subtitle overlay
  --hide-visible-overlay       Force hide visible subtitle overlay
  --show-invisible-overlay     Force show invisible interactive overlay
  --hide-invisible-overlay     Force hide invisible interactive overlay
  --copy-subtitle              Copy current subtitle text
  --copy-subtitle-multiple     Start multi-copy mode
  --mine-sentence              Mine sentence card from current subtitle
  --mine-sentence-multiple     Start multi-mine sentence mode
  --update-last-card-from-clipboard  Update last card from clipboard
  --toggle-secondary-sub       Cycle secondary subtitle mode
  --trigger-field-grouping     Trigger Kiku field grouping
  --trigger-subsync            Run subtitle sync
  --mark-audio-card            Mark last card as audio card
  --open-runtime-options       Open runtime options palette
  --auto-start-overlay  Auto-hide mpv subtitles on connect (show overlay)
  --socket PATH         Override MPV IPC socket/pipe path
  --backend BACKEND     Override window tracker backend (auto, hyprland, sway, x11, macos)
  --port PORT           Texthooker server port (default: ${DEFAULT_TEXTHOOKER_PORT})
  --verbose             Enable debug logging (equivalent to --log-level debug)
  --log-level LEVEL     Set log level: debug, info, warn, error
  --generate-config     Generate default config.jsonc from centralized config registry
  --config-path PATH    Target config path for --generate-config
  --backup-overwrite    With --generate-config, backup and overwrite existing file
  --dev                 Run in development mode
  --debug               Alias for --dev
  --help                Show this help
`);
}

function hasExplicitCommand(args: CliArgs): boolean {
  return (
    args.start ||
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
    args.generateConfig ||
    args.help
  );
}

function shouldStartApp(args: CliArgs): boolean {
  if (args.stop && !args.start) return false;
  if (
    args.start ||
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.toggleInvisibleOverlay ||
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
    args.texthooker
  ) {
    return true;
  }
  return false;
}

function formatBackupTimestamp(date = new Date()): string {
  const pad = (v: number): string => String(v).padStart(2, "0");
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function promptYesNo(question: string): Promise<boolean> {
  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question(question, (answer) => {
      rl.close();
      const normalized = answer.trim().toLowerCase();
      resolve(normalized === "y" || normalized === "yes");
    });
  });
}

async function generateDefaultConfigFile(args: CliArgs): Promise<number> {
  const targetPath = args.configPath
    ? path.resolve(args.configPath)
    : path.join(CONFIG_DIR, "config.jsonc");
  const template = generateConfigTemplate(DEFAULT_CONFIG);

  if (fs.existsSync(targetPath)) {
    if (args.backupOverwrite) {
      const backupPath = `${targetPath}.bak.${formatBackupTimestamp()}`;
      fs.copyFileSync(targetPath, backupPath);
      fs.writeFileSync(targetPath, template, "utf-8");
      console.log(`Backed up existing config to ${backupPath}`);
      console.log(`Generated config at ${targetPath}`);
      return 0;
    }

    if (!process.stdin.isTTY || !process.stdout.isTTY) {
      console.error(
        `Config exists at ${targetPath}. Re-run with --backup-overwrite to back up and overwrite.`,
      );
      return 1;
    }

    const confirmed = await promptYesNo(
      `Config exists at ${targetPath}. Back up and overwrite? [y/N] `,
    );
    if (!confirmed) {
      console.log("Config generation cancelled.");
      return 0;
    }

    const backupPath = `${targetPath}.bak.${formatBackupTimestamp()}`;
    fs.copyFileSync(targetPath, backupPath);
    fs.writeFileSync(targetPath, template, "utf-8");
    console.log(`Backed up existing config to ${backupPath}`);
    console.log(`Generated config at ${targetPath}`);
    return 0;
  }

  const parentDir = path.dirname(targetPath);
  if (!fs.existsSync(parentDir)) {
    fs.mkdirSync(parentDir, { recursive: true });
  }
  fs.writeFileSync(targetPath, template, "utf-8");
  console.log(`Generated config at ${targetPath}`);
  return 0;
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

// Shortcut state tracking
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

function commandNeedsOverlayRuntime(args: CliArgs): boolean {
  return (
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.toggleInvisibleOverlay ||
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
    args.openRuntimeOptions
  );
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

function execCommand(
  command: string,
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    childProcess.exec(command, { timeout: 10000 }, (err, stdout, stderr) => {
      if (err) {
        reject(err);
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

async function resolveJimakuApiKey(): Promise<string | null> {
  const config = getJimakuConfig();
  if (config.apiKey && config.apiKey.trim()) {
    console.log("[jimaku] API key found in config");
    return config.apiKey.trim();
  }
  if (config.apiKeyCommand && config.apiKeyCommand.trim()) {
    try {
      const { stdout } = await execCommand(config.apiKeyCommand);
      const key = stdout.trim();
      console.log(
        `[jimaku] apiKeyCommand result: ${key.length > 0 ? "key obtained" : "empty output"}`,
      );
      return key.length > 0 ? key : null;
    } catch (err) {
      console.error(
        "Failed to run jimaku.apiKeyCommand:",
        (err as Error).message,
      );
      return null;
    }
  }
  console.log(
    "[jimaku] No API key configured (neither apiKey nor apiKeyCommand set)",
  );
  return null;
}

function getRetryAfter(headers: http.IncomingHttpHeaders): number | undefined {
  const value = headers["x-ratelimit-reset-after"];
  if (!value) return undefined;
  const raw = Array.isArray(value) ? value[0] : value;
  const parsed = Number.parseFloat(raw);
  if (!Number.isFinite(parsed)) return undefined;
  return parsed;
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

  const baseUrl = getJimakuBaseUrl();
  const url = new URL(endpoint, baseUrl);
  for (const [key, value] of Object.entries(query)) {
    if (value === null || value === undefined) continue;
    url.searchParams.set(key, String(value));
  }

  console.log(`[jimaku] GET ${url.toString()}`);
  const transport = url.protocol === "https:" ? https : http;

  return new Promise((resolve) => {
    const req = transport.request(
      url,
      {
        method: "GET",
        headers: {
          Authorization: apiKey,
          "User-Agent": "SubMiner",
        },
      },
      (res) => {
        let data = "";
        res.on("data", (chunk) => {
          data += chunk.toString();
        });
        res.on("end", () => {
          const status = res.statusCode || 0;
          console.log(`[jimaku] Response HTTP ${status} for ${endpoint}`);
          if (status >= 200 && status < 300) {
            try {
              const parsed = JSON.parse(data) as T;
              resolve({ ok: true, data: parsed });
            } catch (err) {
              console.error(`[jimaku] JSON parse error: ${data.slice(0, 200)}`);
              resolve({
                ok: false,
                error: { error: "Failed to parse Jimaku response JSON." },
              });
            }
            return;
          }

          let errorMessage = `Jimaku API error (HTTP ${status})`;
          try {
            const parsed = JSON.parse(data) as { error?: string };
            if (parsed && parsed.error) {
              errorMessage = parsed.error;
            }
          } catch (err) {
            // Ignore parse errors.
          }
          console.error(`[jimaku] API error: ${errorMessage}`);

          resolve({
            ok: false,
            error: {
              error: errorMessage,
              code: status || undefined,
              retryAfter:
                status === 429 ? getRetryAfter(res.headers) : undefined,
            },
          });
        });
      },
    );

    req.on("error", (err) => {
      console.error(`[jimaku] Network error: ${(err as Error).message}`);
      resolve({
        ok: false,
        error: { error: `Jimaku request failed: ${(err as Error).message}` },
      });
    });

    req.end();
  });
}

function matchEpisodeFromName(name: string): {
  season: number | null;
  episode: number | null;
  index: number | null;
  confidence: "high" | "medium" | "low";
} {
  const seasonEpisode = name.match(/S(\d{1,2})E(\d{1,3})/i);
  if (seasonEpisode && seasonEpisode.index !== undefined) {
    return {
      season: Number.parseInt(seasonEpisode[1], 10),
      episode: Number.parseInt(seasonEpisode[2], 10),
      index: seasonEpisode.index,
      confidence: "high",
    };
  }

  const alt = name.match(/(\d{1,2})x(\d{1,3})/i);
  if (alt && alt.index !== undefined) {
    return {
      season: Number.parseInt(alt[1], 10),
      episode: Number.parseInt(alt[2], 10),
      index: alt.index,
      confidence: "high",
    };
  }

  const epOnly = name.match(/(?:^|[\s._-])E(?:P)?(\d{1,3})(?:\b|[\s._-])/i);
  if (epOnly && epOnly.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(epOnly[1], 10),
      index: epOnly.index,
      confidence: "medium",
    };
  }

  const numeric = name.match(/(?:^|[-–—]\s*)(\d{1,3})\s*[-–—]/);
  if (numeric && numeric.index !== undefined) {
    return {
      season: null,
      episode: Number.parseInt(numeric[1], 10),
      index: numeric.index,
      confidence: "medium",
    };
  }

  return { season: null, episode: null, index: null, confidence: "low" };
}

function detectSeasonFromDir(mediaPath: string): number | null {
  const parent = path.basename(path.dirname(mediaPath));
  const match = parent.match(/(?:Season|S)\s*(\d{1,2})/i);
  if (!match) return null;
  const parsed = Number.parseInt(match[1], 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function cleanupTitle(value: string): string {
  return value
    .replace(/^[\s-–—]+/, "")
    .replace(/[\s-–—]+$/, "")
    .replace(/\s+/g, " ")
    .trim();
}

function parseMediaInfo(mediaPath: string | null): JimakuMediaInfo {
  if (!mediaPath) {
    return {
      title: "",
      season: null,
      episode: null,
      confidence: "low",
      filename: "",
      rawTitle: "",
    };
  }

  const filename = path.basename(mediaPath);
  let name = filename.replace(/\.[^/.]+$/, "");
  name = name.replace(/\[[^\]]*]/g, " ");
  name = name.replace(/\(\d{4}\)/g, " ");
  name = name.replace(/[._]/g, " ");
  name = name.replace(/[–—]/g, "-");
  name = name.replace(/\s+/g, " ").trim();

  const parsed = matchEpisodeFromName(name);
  let titlePart = name;
  if (parsed.index !== null) {
    titlePart = name.slice(0, parsed.index);
  }

  const seasonFromDir = parsed.season ?? detectSeasonFromDir(mediaPath);
  const title = cleanupTitle(titlePart || name);

  return {
    title,
    season: seasonFromDir,
    episode: parsed.episode,
    confidence: parsed.confidence,
    filename,
    rawTitle: name,
  };
}

function getSubtitlePositionFilePath(mediaPath: string): string {
  const key = normalizeMediaPathForSubtitlePosition(mediaPath);
  const hash = crypto.createHash("sha256").update(key).digest("hex");
  return path.join(SUBTITLE_POSITIONS_DIR, `${hash}.json`);
}

function normalizeMediaPathForSubtitlePosition(mediaPath: string): string {
  const trimmed = mediaPath.trim();
  if (!trimmed) return trimmed;

  // Keep URL-like targets as-is; local files are normalized to stable absolute paths.
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

function getTexthookerPath(): string | null {
  const searchPaths = [
    path.join(__dirname, "..", "vendor", "texthooker-ui", "docs"),
    path.join(process.resourcesPath, "app", "vendor", "texthooker-ui", "docs"),
  ];
  for (const p of searchPaths) {
    if (fs.existsSync(path.join(p, "index.html"))) {
      return p;
    }
  }
  return null;
}

function startTexthookerServer(port: number): http.Server | null {
  const texthookerPath = getTexthookerPath();
  if (!texthookerPath) {
    console.error("texthooker-ui not found");
    return null;
  }

  texthookerServer = http.createServer((req, res) => {
    let urlPath = (req.url || "/").split("?")[0];
    let filePath = path.join(
      texthookerPath,
      urlPath === "/" ? "index.html" : urlPath,
    );

    const ext = path.extname(filePath);
    const mimeTypes: Record<string, string> = {
      ".html": "text/html",
      ".js": "application/javascript",
      ".css": "text/css",
      ".json": "application/json",
      ".png": "image/png",
      ".svg": "image/svg+xml",
      ".ttf": "font/ttf",
      ".woff": "font/woff",
      ".woff2": "font/woff2",
    };

    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        res.end("Not found");
        return;
      }
      res.writeHead(200, { "Content-Type": mimeTypes[ext] || "text/plain" });
      res.end(data);
    });
  });

  texthookerServer.listen(port, "127.0.0.1", () => {
    console.log(`Texthooker server running at http://127.0.0.1:${port}`);
  });

  return texthookerServer;
}

function stopTexthookerServer(): void {
  if (texthookerServer) {
    texthookerServer.close();
    texthookerServer = null;
  }
}

function hasMpvWebsocket(): boolean {
  const mpvWebsocketPath = path.join(
    os.homedir(),
    ".config",
    "mpv",
    "mpv_websocket",
  );
  return fs.existsSync(mpvWebsocketPath);
}

function startSubtitleWebSocketServer(port: number): void {
  subtitleWebSocketServer = new WebSocket.Server({ port, host: "127.0.0.1" });

  subtitleWebSocketServer.on("connection", (ws: WebSocket) => {
    console.log("WebSocket client connected");
    if (currentSubText) {
      ws.send(JSON.stringify({ sentence: currentSubText }));
    }
  });

  subtitleWebSocketServer.on("error", (err: Error) => {
    console.error("WebSocket server error:", err.message);
  });

  console.log(`Subtitle WebSocket server running on ws://127.0.0.1:${port}`);
}

function broadcastSubtitle(text: string): void {
  if (!subtitleWebSocketServer) return;
  const message = JSON.stringify({ sentence: text });
  for (const client of subtitleWebSocketServer.clients) {
    if (client.readyState === WebSocket.OPEN) {
      client.send(message);
    }
  }
}

function stopSubtitleWebSocketServer(): void {
  if (subtitleWebSocketServer) {
    subtitleWebSocketServer.close();
    subtitleWebSocketServer = null;
  }
}

interface MpvTrack {
  id?: number;
  type?: string;
  selected?: boolean;
  external?: boolean;
  lang?: string;
  title?: string;
  codec?: string;
  "ff-index"?: number;
  "external-filename"?: string;
}

interface SubsyncResolvedConfig {
  defaultMode: SubsyncMode;
  alassPath: string;
  ffsubsyncPath: string;
  ffmpegPath: string;
}

const DEFAULT_SUBSYNC_EXECUTABLE_PATHS = {
  alass: "/usr/bin/alass",
  ffsubsync: "/usr/bin/ffsubsync",
  ffmpeg: "/usr/bin/ffmpeg",
} as const;

interface SubsyncContext {
  videoPath: string;
  primaryTrack: MpvTrack;
  secondaryTrack: MpvTrack | null;
  sourceTracks: MpvTrack[];
  audioStreamIndex: number | null;
}

interface CommandResult {
  ok: boolean;
  code: number | null;
  stderr: string;
  stdout: string;
  error?: string;
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

interface FileExtractionResult {
  path: string;
  temporary: boolean;
}

function getSubsyncConfig(
  config: SubsyncConfig | undefined,
): SubsyncResolvedConfig {
  const resolvePath = (value: string | undefined, fallback: string): string => {
    const trimmed = value?.trim();
    return trimmed && trimmed.length > 0 ? trimmed : fallback;
  };

  return {
    defaultMode: config?.defaultMode ?? DEFAULT_CONFIG.subsync.defaultMode,
    alassPath: resolvePath(
      config?.alass_path,
      DEFAULT_SUBSYNC_EXECUTABLE_PATHS.alass,
    ),
    ffsubsyncPath: resolvePath(
      config?.ffsubsync_path,
      DEFAULT_SUBSYNC_EXECUTABLE_PATHS.ffsubsync,
    ),
    ffmpegPath: resolvePath(
      config?.ffmpeg_path,
      DEFAULT_SUBSYNC_EXECUTABLE_PATHS.ffmpeg,
    ),
  };
}

function hasPathSeparators(value: string): boolean {
  return value.includes("/") || value.includes("\\");
}

function fileExists(pathOrEmpty: string): boolean {
  if (!pathOrEmpty) return false;
  try {
    return fs.existsSync(pathOrEmpty);
  } catch {
    return false;
  }
}

function formatTrackLabel(track: MpvTrack): string {
  const trackId = typeof track.id === "number" ? track.id : -1;
  const source = track.external ? "External" : "Internal";
  const lang = track.lang || track.title || "unknown";
  const active = track.selected ? " (active)" : "";
  return `${source} #${trackId} - ${lang}${active}`;
}

function getTrackById(
  tracks: MpvTrack[],
  trackId: number | null,
): MpvTrack | null {
  if (trackId === null) return null;
  return tracks.find((track) => track.id === trackId) ?? null;
}

function codecToExtension(codec: string | undefined): string | null {
  if (!codec) return null;
  const normalized = codec.toLowerCase();
  if (normalized === "subrip" || normalized === "srt") return "srt";
  if (normalized === "ass" || normalized === "ssa") return "ass";
  return null;
}

function runCommand(
  executable: string,
  args: string[],
  timeoutMs = 120000,
): Promise<CommandResult> {
  return new Promise((resolve) => {
    const child = childProcess.spawn(executable, args, {
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    const timeout = setTimeout(() => {
      child.kill("SIGKILL");
    }, timeoutMs);

    child.stdout.on("data", (chunk: Buffer) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk: Buffer) => {
      stderr += chunk.toString();
    });
    child.on("error", (error: Error) => {
      clearTimeout(timeout);
      resolve({
        ok: false,
        code: null,
        stderr,
        stdout,
        error: error.message,
      });
    });
    child.on("close", (code: number | null) => {
      clearTimeout(timeout);
      resolve({
        ok: code === 0,
        code,
        stderr,
        stdout,
      });
    });
  });
}

function loadKeybindings(): Keybinding[] {
  const config = getResolvedConfig();
  const userBindings = config.keybindings || [];

  const bindingMap = new Map<string, (string | number)[] | null>();

  for (const binding of DEFAULT_KEYBINDINGS) {
    bindingMap.set(binding.key, binding.command);
  }

  for (const binding of userBindings) {
    if (binding.command === null) {
      bindingMap.delete(binding.key);
    } else {
      bindingMap.set(binding.key, binding.command);
    }
  }

  keybindings = [];
  for (const [key, command] of bindingMap) {
    if (command !== null) {
      keybindings.push({ key, command });
    }
  }

  return keybindings;
}

const initialArgs = parseArgs(process.argv);
if (initialArgs.logLevel) {
  process.env.SUBMINER_LOG_LEVEL = initialArgs.logLevel;
} else if (initialArgs.verbose) {
  process.env.SUBMINER_LOG_LEVEL = "debug";
}

function getElectronOzonePlatformHint(): string | null {
  const hint = process.env.ELECTRON_OZONE_PLATFORM_HINT?.trim().toLowerCase();
  if (hint) return hint;
  const ozone = process.env.OZONE_PLATFORM?.trim().toLowerCase();
  if (ozone) return ozone;
  return null;
}

function enforceUnsupportedWaylandMode(args: CliArgs): void {
  if (process.platform !== "linux") return;
  if (!shouldStartApp(args)) return;
  const hint = getElectronOzonePlatformHint();
  if (hint !== "wayland") return;

  const message =
    "Unsupported Electron backend: Wayland. Set ELECTRON_OZONE_PLATFORM_HINT=x11 and restart SubMiner.";
  console.error(message);
  throw new Error(message);
}

enforceUnsupportedWaylandMode(initialArgs);

let mpvSocketPath = initialArgs.socketPath ?? getDefaultSocketPath();
let texthookerPort = initialArgs.texthookerPort ?? DEFAULT_TEXTHOOKER_PORT;
const backendOverride = initialArgs.backend ?? null;
const autoStartOverlay = initialArgs.autoStartOverlay;
const texthookerOnlyMode = initialArgs.texthooker;

if (initialArgs.generateConfig && !shouldStartApp(initialArgs)) {
  generateDefaultConfigFile(initialArgs)
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
      printHelp();
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
        loadKeybindings();

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
          (wsEnabled === "auto" && !hasMpvWebsocket())
        ) {
          startSubtitleWebSocketServer(wsPort);
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
        stopSubtitleWebSocketServer();
        stopTexthookerServer();
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
    if (texthookerServer) {
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
    if (!texthookerServer) {
      startTexthookerServer(texthookerPort);
    }
    const config = getResolvedConfig();
    const openBrowser = config.texthooker?.openBrowser !== false;
    if (openBrowser) {
      shell.openExternal(`http://127.0.0.1:${texthookerPort}`);
    }
    console.log(`Texthooker available at http://127.0.0.1:${texthookerPort}`);
  } else if (args.help) {
    printHelp();
    if (!mainWindow) app.quit();
  }
}

function handleInitialArgs(): void {
  handleCliCommand(initialArgs, "initial");
}

function asFiniteNumber(
  value: unknown,
  fallback: number,
  min?: number,
  max?: number,
): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return fallback;
  }
  let next = value;
  if (typeof min === "number") next = Math.max(min, next);
  if (typeof max === "number") next = Math.min(max, next);
  return next;
}

function asString(value: unknown, fallback: string): string {
  if (typeof value !== "string") return fallback;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

function asBoolean(value: unknown, fallback: boolean): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "yes" || normalized === "true" || normalized === "1") {
      return true;
    }
    if (normalized === "no" || normalized === "false" || normalized === "0") {
      return false;
    }
  }
  return fallback;
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
        broadcastSubtitle(currentSubText);
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
        broadcastSubtitle(currentSubText);
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
  if (!text || !mecabTokenizer) {
    return { text, tokens: null };
  }

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
  // Copy extension to writable location on Linux and macOS
  // MV3 service workers need write access for IndexedDB/storage
  // App bundles on macOS are read-only, causing service worker failures
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

  console.log("Yomitan search paths:", searchPaths);
  console.log("Found Yomitan at:", extPath);

  if (!extPath) {
    console.error("Yomitan extension not found in any search path");
    console.error("Install Yomitan to one of:", searchPaths);
    return null;
  }

  extPath = ensureExtensionCopy(extPath);
  console.log("Using extension path:", extPath);

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
    console.log("Yomitan extension loaded successfully:", yomitanExt.id);
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
  console.log(`Loading ${kind} overlay HTML from:`, htmlPath);
  console.log("HTML file exists:", fs.existsSync(htmlPath));

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
    console.log(`${kind} overlay HTML loaded successfully`);
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
      console.log("MPV window found:", geometry);
      updateOverlayBounds(geometry);
      if (visibleOverlayVisible) {
        updateVisibleOverlayVisibility();
      }
      if (invisibleOverlayVisible) {
        updateInvisibleOverlayVisibility();
      }
    };
    windowTracker.onWindowLost = () => {
      console.log("MPV window lost");
      for (const window of getOverlayWindows()) {
        window.hide();
      }
      // Keep overlay shortcuts registered; tracking loss can be transient.
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
  console.log("openYomitanSettings called");

  if (!yomitanExt) {
    console.error("Yomitan extension not loaded - yomitanExt is:", yomitanExt);
    console.error(
      "This may be due to Manifest V3 service worker issues with Electron",
    );
    return;
  }

  if (yomitanSettingsWindow && !yomitanSettingsWindow.isDestroyed()) {
    console.log("Settings window already exists, focusing");
    yomitanSettingsWindow.focus();
    return;
  }

  console.log("Creating new settings window for extension:", yomitanExt.id);

  yomitanSettingsWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      session: session.defaultSession,
    },
  });

  const settingsUrl = `chrome-extension://${yomitanExt.id}/settings.html`;
  console.log("Loading settings URL:", settingsUrl);

  let loadAttempts = 0;
  const maxAttempts = 3;

  function attemptLoad(): void {
    yomitanSettingsWindow!
      .loadURL(settingsUrl)
      .then(() => {
        console.log("Settings URL loaded successfully");
      })
      .catch((err: Error) => {
        console.error("Failed to load settings URL:", err);
        loadAttempts++;
        if (
          loadAttempts < maxAttempts &&
          yomitanSettingsWindow &&
          !yomitanSettingsWindow.isDestroyed()
        ) {
          console.log(
            `Retrying in 500ms (attempt ${loadAttempts + 1}/${maxAttempts})`,
          );
          setTimeout(attemptLoad, 500);
        }
      });
  }

  attemptLoad();

  yomitanSettingsWindow.webContents.on(
    "did-fail-load",
    (_event, errorCode, errorDescription) => {
      console.error(
        "Settings page failed to load:",
        errorCode,
        errorDescription,
      );
    },
  );

  yomitanSettingsWindow.webContents.on("did-finish-load", () => {
    console.log("Settings page loaded successfully");
  });

  setTimeout(() => {
    if (yomitanSettingsWindow && !yomitanSettingsWindow.isDestroyed()) {
      yomitanSettingsWindow.setSize(
        yomitanSettingsWindow.getSize()[0],
        yomitanSettingsWindow.getSize()[1],
      );
      yomitanSettingsWindow.webContents.invalidate();
      yomitanSettingsWindow.show();
    }
  }, 500);

  yomitanSettingsWindow.on("closed", () => {
    yomitanSettingsWindow = null;
  });
}

function registerGlobalShortcuts(): void {
  const shortcuts = getConfiguredShortcuts();
  const visibleShortcut = shortcuts.toggleVisibleOverlayGlobal;
  const invisibleShortcut = shortcuts.toggleInvisibleOverlayGlobal;
  const normalizedVisible = visibleShortcut?.replace(/\s+/g, "").toLowerCase();
  const normalizedInvisible = invisibleShortcut
    ?.replace(/\s+/g, "")
    .toLowerCase();

  if (visibleShortcut) {
    const toggleVisibleRegistered = globalShortcut.register(
      visibleShortcut,
      () => {
        toggleVisibleOverlay();
      },
    );
    if (!toggleVisibleRegistered) {
      console.warn(
        `Failed to register global shortcut toggleVisibleOverlayGlobal: ${visibleShortcut}`,
      );
    }
  }

  if (
    invisibleShortcut &&
    normalizedInvisible &&
    normalizedInvisible !== normalizedVisible
  ) {
    const toggleInvisibleRegistered = globalShortcut.register(
      invisibleShortcut,
      () => {
        toggleInvisibleOverlay();
      },
    );
    if (!toggleInvisibleRegistered) {
      console.warn(
        `Failed to register global shortcut toggleInvisibleOverlayGlobal: ${invisibleShortcut}`,
      );
    }
  } else if (
    invisibleShortcut &&
    normalizedInvisible &&
    normalizedInvisible === normalizedVisible
  ) {
    console.warn(
      "Skipped registering toggleInvisibleOverlayGlobal because it collides with toggleVisibleOverlayGlobal",
    );
  }

  const settingsRegistered = globalShortcut.register("Alt+Shift+Y", () => {
    openYomitanSettings();
  });
  if (!settingsRegistered) {
    console.warn("Failed to register global shortcut: Alt+Shift+Y");
  }

  if (isDev) {
    const devtoolsRegistered = globalShortcut.register("F12", () => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.toggleDevTools();
      }
    });
    if (!devtoolsRegistered) {
      console.warn("Failed to register global shortcut: F12");
    }
  }
}

function getConfiguredShortcuts() {
  const config = getResolvedConfig();
  const normalizeShortcut = (
    value: string | null | undefined,
  ): string | null | undefined => {
    if (typeof value !== "string") return value;
    return value
      .replace(/\bKey([A-Z])\b/g, "$1")
      .replace(/\bDigit([0-9])\b/g, "$1");
  };

  return {
    toggleVisibleOverlayGlobal: normalizeShortcut(
      config.shortcuts?.toggleVisibleOverlayGlobal ??
        DEFAULT_CONFIG.shortcuts.toggleVisibleOverlayGlobal,
    ),
    toggleInvisibleOverlayGlobal: normalizeShortcut(
      config.shortcuts?.toggleInvisibleOverlayGlobal ??
        DEFAULT_CONFIG.shortcuts.toggleInvisibleOverlayGlobal,
    ),
    copySubtitle: normalizeShortcut(
      config.shortcuts?.copySubtitle ?? DEFAULT_CONFIG.shortcuts.copySubtitle,
    ),
    copySubtitleMultiple: normalizeShortcut(
      config.shortcuts?.copySubtitleMultiple ??
        DEFAULT_CONFIG.shortcuts.copySubtitleMultiple,
    ),
    updateLastCardFromClipboard: normalizeShortcut(
      config.shortcuts?.updateLastCardFromClipboard ??
        DEFAULT_CONFIG.shortcuts.updateLastCardFromClipboard,
    ),
    triggerFieldGrouping: normalizeShortcut(
      config.shortcuts?.triggerFieldGrouping ??
        DEFAULT_CONFIG.shortcuts.triggerFieldGrouping,
    ),
    triggerSubsync: normalizeShortcut(
      config.shortcuts?.triggerSubsync ??
        DEFAULT_CONFIG.shortcuts.triggerSubsync,
    ),
    mineSentence: normalizeShortcut(
      config.shortcuts?.mineSentence ?? DEFAULT_CONFIG.shortcuts.mineSentence,
    ),
    mineSentenceMultiple: normalizeShortcut(
      config.shortcuts?.mineSentenceMultiple ??
        DEFAULT_CONFIG.shortcuts.mineSentenceMultiple,
    ),
    multiCopyTimeoutMs:
      config.shortcuts?.multiCopyTimeoutMs ??
      DEFAULT_CONFIG.shortcuts.multiCopyTimeoutMs,
    toggleSecondarySub: normalizeShortcut(
      config.shortcuts?.toggleSecondarySub ??
        DEFAULT_CONFIG.shortcuts.toggleSecondarySub,
    ),
    markAudioCard: normalizeShortcut(
      config.shortcuts?.markAudioCard ?? DEFAULT_CONFIG.shortcuts.markAudioCard,
    ),
    openRuntimeOptions: normalizeShortcut(
      config.shortcuts?.openRuntimeOptions ??
        DEFAULT_CONFIG.shortcuts.openRuntimeOptions,
    ),
  };
}

function shouldUseMarkAudioCardLocalFallback(input: Electron.Input): boolean {
  const shortcuts = getConfiguredShortcuts();
  if (!shortcuts.markAudioCard) return false;
  if (globalShortcut.isRegistered(shortcuts.markAudioCard)) return false;

  const normalized = shortcuts.markAudioCard.replace(/\s+/g, "").toLowerCase();
  const supportsFallback =
    normalized === "commandorcontrol+shift+a" ||
    normalized === "cmdorctrl+shift+a" ||
    normalized === "control+shift+a" ||
    normalized === "ctrl+shift+a";
  if (!supportsFallback) return false;

  if (input.type !== "keyDown" || input.isAutoRepeat) return false;
  if ((input.key || "").toLowerCase() !== "a") return false;
  if (!input.shift || input.alt) return false;

  if (process.platform === "darwin") {
    return Boolean(input.meta || input.control);
  }
  return Boolean(input.control);
}

function shouldUseRuntimeOptionsLocalFallback(input: Electron.Input): boolean {
  const shortcuts = getConfiguredShortcuts();
  if (!shortcuts.openRuntimeOptions) return false;
  if (globalShortcut.isRegistered(shortcuts.openRuntimeOptions)) return false;

  const normalized = shortcuts.openRuntimeOptions
    .replace(/\s+/g, "")
    .toLowerCase();
  const supportsFallback =
    normalized === "commandorcontrol+shift+o" ||
    normalized === "cmdorctrl+shift+o" ||
    normalized === "control+shift+o" ||
    normalized === "ctrl+shift+o";
  if (!supportsFallback) return false;

  if (input.type !== "keyDown" || input.isAutoRepeat) return false;
  if ((input.key || "").toLowerCase() !== "o") return false;
  if (!input.shift || input.alt) return false;

  if (process.platform === "darwin") {
    return Boolean(input.meta || input.control);
  }
  return Boolean(input.control);
}

function isGlobalShortcutRegisteredSafe(accelerator: string): boolean {
  try {
    return globalShortcut.isRegistered(accelerator);
  } catch {
    return false;
  }
}

function shortcutMatchesInputForLocalFallback(
  input: Electron.Input,
  accelerator: string,
  allowWhenRegistered = false,
): boolean {
  if (input.type !== "keyDown" || input.isAutoRepeat) return false;
  if (!accelerator) return false;
  if (!allowWhenRegistered && isGlobalShortcutRegisteredSafe(accelerator)) {
    return false;
  }

  const normalized = accelerator
    .replace(/\s+/g, "")
    .replace(/cmdorctrl/gi, "CommandOrControl")
    .toLowerCase();
  const parts = normalized.split("+").filter(Boolean);
  if (parts.length === 0) return false;

  const keyToken = parts[parts.length - 1];
  const modifierTokens = new Set(parts.slice(0, -1));
  const allowedModifiers = new Set([
    "shift",
    "alt",
    "meta",
    "control",
    "commandorcontrol",
  ]);
  for (const token of modifierTokens) {
    if (!allowedModifiers.has(token)) return false;
  }

  const inputKey = (input.key || "").toLowerCase();
  if (keyToken.length === 1) {
    if (inputKey !== keyToken) return false;
  } else if (keyToken.startsWith("key") && keyToken.length === 4) {
    if (inputKey !== keyToken.slice(3)) return false;
  } else {
    return false;
  }

  const expectedShift = modifierTokens.has("shift");
  const expectedAlt = modifierTokens.has("alt");
  const expectedMeta = modifierTokens.has("meta");
  const expectedControl = modifierTokens.has("control");
  const expectedCommandOrControl = modifierTokens.has("commandorcontrol");

  if (Boolean(input.shift) !== expectedShift) return false;
  if (Boolean(input.alt) !== expectedAlt) return false;

  if (expectedCommandOrControl) {
    const hasCmdOrCtrl =
      process.platform === "darwin"
        ? Boolean(input.meta || input.control)
        : Boolean(input.control);
    if (!hasCmdOrCtrl) return false;
  } else {
    if (process.platform === "darwin") {
      if (input.meta || input.control) return false;
    } else if (input.control) {
      return false;
    }
  }

  if (expectedMeta && !input.meta) return false;
  if (!expectedMeta && modifierTokens.has("meta") === false && input.meta) {
    if (!expectedCommandOrControl) return false;
  }

  if (expectedControl && !input.control) return false;
  if (
    !expectedControl &&
    modifierTokens.has("control") === false &&
    input.control
  ) {
    if (!expectedCommandOrControl) return false;
  }

  return true;
}

function tryHandleOverlayShortcutLocalFallback(input: Electron.Input): boolean {
  const shortcuts = getConfiguredShortcuts();
  const handlers: Array<{
    accelerator: string | null | undefined;
    run: () => void;
    allowWhenRegistered?: boolean;
  }> = [
    {
      accelerator: shortcuts.openRuntimeOptions,
      run: () => {
        openRuntimeOptionsPalette();
      },
    },
    {
      accelerator: shortcuts.markAudioCard,
      run: () => {
        markLastCardAsAudioCard().catch((err) => {
          console.error("markLastCardAsAudioCard failed:", err);
          showMpvOsd(`Audio card failed: ${(err as Error).message}`);
        });
      },
    },
    {
      accelerator: shortcuts.copySubtitleMultiple,
      run: () => {
        startPendingMultiCopy(shortcuts.multiCopyTimeoutMs);
      },
    },
    {
      accelerator: shortcuts.copySubtitle,
      run: () => {
        copyCurrentSubtitle();
      },
    },
    {
      accelerator: shortcuts.toggleSecondarySub,
      run: () => cycleSecondarySubMode(),
      allowWhenRegistered: true,
    },
    {
      accelerator: shortcuts.updateLastCardFromClipboard,
      run: () => {
        updateLastCardFromClipboard().catch((err) => {
          console.error("updateLastCardFromClipboard failed:", err);
          showMpvOsd(`Update failed: ${(err as Error).message}`);
        });
      },
    },
    {
      accelerator: shortcuts.triggerFieldGrouping,
      run: () => {
        triggerFieldGrouping().catch((err) => {
          console.error("triggerFieldGrouping failed:", err);
          showMpvOsd(`Field grouping failed: ${(err as Error).message}`);
        });
      },
    },
    {
      accelerator: shortcuts.triggerSubsync,
      run: () => {
        triggerSubsyncFromConfig().catch((err) => {
          console.error("triggerSubsyncFromConfig failed:", err);
          showMpvOsd(`Subsync failed: ${(err as Error).message}`);
        });
      },
    },
    {
      accelerator: shortcuts.mineSentence,
      run: () => {
        mineSentenceCard().catch((err) => {
          console.error("mineSentenceCard failed:", err);
          showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
        });
      },
    },
    {
      accelerator: shortcuts.mineSentenceMultiple,
      run: () => {
        startPendingMineSentenceMultiple(shortcuts.multiCopyTimeoutMs);
      },
    },
  ];

  for (const handler of handlers) {
    if (!handler.accelerator) continue;
    if (
      shortcutMatchesInputForLocalFallback(
        input,
        handler.accelerator,
        handler.allowWhenRegistered === true,
      )
    ) {
      handler.run();
      return true;
    }
  }

  return false;
}

function cycleSecondarySubMode(): void {
  // Some platforms can trigger both global and in-window handlers for one key press.
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

function getMpvClientForSubsync(): MpvIpcClient {
  if (!mpvClient || !mpvClient.connected) {
    throw new Error("MPV not connected");
  }
  return mpvClient as MpvIpcClient;
}

async function gatherSubsyncContext(
  client: MpvIpcClient,
): Promise<SubsyncContext> {
  const [videoPathRaw, sidRaw, secondarySidRaw, trackListRaw] =
    await Promise.all([
      client.requestProperty("path"),
      client.requestProperty("sid"),
      client.requestProperty("secondary-sid"),
      client.requestProperty("track-list"),
    ]);

  const videoPath = typeof videoPathRaw === "string" ? videoPathRaw : "";
  if (!videoPath) {
    throw new Error("No video is currently loaded");
  }

  const tracks = Array.isArray(trackListRaw)
    ? (trackListRaw as MpvTrack[])
    : [];
  const subtitleTracks = tracks.filter((track) => track.type === "sub");
  const sid = typeof sidRaw === "number" ? sidRaw : null;
  const secondarySid =
    typeof secondarySidRaw === "number" ? secondarySidRaw : null;

  const primaryTrack = subtitleTracks.find((track) => track.id === sid);
  if (!primaryTrack) {
    throw new Error("No active subtitle track found");
  }

  const secondaryTrack =
    subtitleTracks.find((track) => track.id === secondarySid) ?? null;
  const sourceTracks = subtitleTracks
    .filter((track) => track.id !== sid)
    .filter((track) => {
      if (!track.external) return true;
      const filename = track["external-filename"];
      return typeof filename === "string" && filename.length > 0;
    });

  return {
    videoPath,
    primaryTrack,
    secondaryTrack,
    sourceTracks,
    audioStreamIndex: client.currentAudioStreamIndex,
  };
}

function ensureExecutablePath(pathOrName: string, name: string): string {
  if (!pathOrName) {
    throw new Error(`Missing ${name} path in config`);
  }

  if (hasPathSeparators(pathOrName) && !fileExists(pathOrName)) {
    throw new Error(`Configured ${name} executable not found: ${pathOrName}`);
  }
  return pathOrName;
}

async function extractSubtitleTrackToFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvTrack,
): Promise<FileExtractionResult> {
  if (track.external) {
    const externalPath = track["external-filename"];
    if (typeof externalPath !== "string" || externalPath.length === 0) {
      throw new Error("External subtitle track has no file path");
    }
    if (!fileExists(externalPath)) {
      throw new Error(`Subtitle file not found: ${externalPath}`);
    }
    return { path: externalPath, temporary: false };
  }

  const ffIndex = track["ff-index"];
  const extension = codecToExtension(track.codec);
  if (
    typeof ffIndex !== "number" ||
    !Number.isInteger(ffIndex) ||
    ffIndex < 0
  ) {
    throw new Error("Internal subtitle track has no valid ff-index");
  }
  if (!extension) {
    throw new Error(`Unsupported subtitle codec: ${track.codec ?? "unknown"}`);
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "subminer-subsync-"));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
  const extraction = await runCommand(ffmpegPath, [
    "-hide_banner",
    "-nostdin",
    "-y",
    "-loglevel",
    "quiet",
    "-an",
    "-vn",
    "-i",
    videoPath,
    "-map",
    `0:${ffIndex}`,
    "-f",
    extension,
    outputPath,
  ]);

  if (!extraction.ok || !fileExists(outputPath)) {
    throw new Error("Failed to extract internal subtitle track with ffmpeg");
  }

  return { path: outputPath, temporary: true };
}

function cleanupTemporaryFile(extraction: FileExtractionResult): void {
  if (!extraction.temporary) return;
  try {
    if (fileExists(extraction.path)) {
      fs.unlinkSync(extraction.path);
    }
  } catch {
    // Ignore cleanup failures
  }
  try {
    const dir = path.dirname(extraction.path);
    if (fs.existsSync(dir)) {
      fs.rmdirSync(dir);
    }
  } catch {
    // Ignore cleanup failures
  }
}

function buildRetimedPath(subPath: string): string {
  const parsed = path.parse(subPath);
  const suffix = `_retimed_${Date.now()}`;
  return path.join(
    parsed.dir,
    `${parsed.name}${suffix}${parsed.ext || ".srt"}`,
  );
}

async function runAlassSync(
  alassPath: string,
  referenceFile: string,
  inputSubtitlePath: string,
  outputPath: string,
): Promise<CommandResult> {
  return runCommand(alassPath, [referenceFile, inputSubtitlePath, outputPath]);
}

async function runFfsubsyncSync(
  ffsubsyncPath: string,
  videoPath: string,
  inputSubtitlePath: string,
  outputPath: string,
  audioStreamIndex: number | null,
): Promise<CommandResult> {
  const args = [videoPath, "-i", inputSubtitlePath, "-o", outputPath];
  if (audioStreamIndex !== null) {
    args.push("--reference-stream", `0:${audioStreamIndex}`);
  }
  return runCommand(ffsubsyncPath, args);
}

function loadSyncedSubtitle(pathToLoad: string): void {
  if (!mpvClient || !mpvClient.connected) {
    throw new Error("MPV disconnected while loading subtitle");
  }
  mpvClient.send({ command: ["sub_add", pathToLoad] });
  mpvClient.send({ command: ["set_property", "sub-delay", 0] });
}

async function subsyncToReference(
  engine: "alass" | "ffsubsync",
  referenceFilePath: string,
  context: SubsyncContext,
  resolved: SubsyncResolvedConfig,
): Promise<SubsyncResult> {
  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, "ffmpeg");
  const primaryExtraction = await extractSubtitleTrackToFile(
    ffmpegPath,
    context.videoPath,
    context.primaryTrack,
  );
  const outputPath = buildRetimedPath(primaryExtraction.path);

  try {
    let result: CommandResult;
    if (engine === "alass") {
      const alassPath = ensureExecutablePath(resolved.alassPath, "alass");
      result = await runAlassSync(
        alassPath,
        referenceFilePath,
        primaryExtraction.path,
        outputPath,
      );
    } else {
      const ffsubsyncPath = ensureExecutablePath(
        resolved.ffsubsyncPath,
        "ffsubsync",
      );
      result = await runFfsubsyncSync(
        ffsubsyncPath,
        context.videoPath,
        primaryExtraction.path,
        outputPath,
        context.audioStreamIndex,
      );
    }

    if (!result.ok || !fileExists(outputPath)) {
      return {
        ok: false,
        message: `${engine} synchronization failed`,
      };
    }

    loadSyncedSubtitle(outputPath);
    return {
      ok: true,
      message: `Subtitle synchronized with ${engine}`,
    };
  } finally {
    cleanupTemporaryFile(primaryExtraction);
  }
}

async function runSubsyncAuto(): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync();
  const context = await gatherSubsyncContext(client);
  const resolved = getSubsyncConfig(getResolvedConfig().subsync);
  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, "ffmpeg");

  if (context.secondaryTrack) {
    let secondaryExtraction: FileExtractionResult | null = null;
    try {
      secondaryExtraction = await extractSubtitleTrackToFile(
        ffmpegPath,
        context.videoPath,
        context.secondaryTrack,
      );
      const alassResult = await subsyncToReference(
        "alass",
        secondaryExtraction.path,
        context,
        resolved,
      );
      if (alassResult.ok) {
        return alassResult;
      }
    } catch (error) {
      // Fall through to ffsubsync fallback
      console.warn("Auto alass sync failed, trying ffsubsync fallback:", error);
    } finally {
      if (secondaryExtraction) {
        cleanupTemporaryFile(secondaryExtraction);
      }
    }
  }

  const ffsubsyncPath = ensureExecutablePath(
    resolved.ffsubsyncPath,
    "ffsubsync",
  );
  if (!ffsubsyncPath) {
    return {
      ok: false,
      message: "No secondary subtitle for alass and ffsubsync not configured",
    };
  }
  return subsyncToReference("ffsubsync", context.videoPath, context, resolved);
}

async function openSubsyncManualPicker(): Promise<void> {
  const client = getMpvClientForSubsync();
  const context = await gatherSubsyncContext(client);
  const payload: SubsyncManualPayload = {
    sourceTracks: context.sourceTracks
      .filter((track) => typeof track.id === "number")
      .map((track) => ({
        id: track.id as number,
        label: formatTrackLabel(track),
      })),
  };
  sendToVisibleOverlay("subsync:open-manual", payload, {
    restoreOnModalClose: "subsync",
  });
}

async function runSubsyncManual(
  request: SubsyncManualRunRequest,
): Promise<SubsyncResult> {
  const client = getMpvClientForSubsync();
  const context = await gatherSubsyncContext(client);
  const resolved = getSubsyncConfig(getResolvedConfig().subsync);

  if (request.engine === "ffsubsync") {
    return subsyncToReference(
      "ffsubsync",
      context.videoPath,
      context,
      resolved,
    );
  }

  const sourceTrack = getTrackById(
    context.sourceTracks,
    request.sourceTrackId ?? null,
  );
  if (!sourceTrack) {
    return { ok: false, message: "Select a subtitle source track for alass" };
  }

  const ffmpegPath = ensureExecutablePath(resolved.ffmpegPath, "ffmpeg");
  let sourceExtraction: FileExtractionResult | null = null;
  try {
    sourceExtraction = await extractSubtitleTrackToFile(
      ffmpegPath,
      context.videoPath,
      sourceTrack,
    );
    return subsyncToReference(
      "alass",
      sourceExtraction.path,
      context,
      resolved,
    );
  } finally {
    if (sourceExtraction) {
      cleanupTemporaryFile(sourceExtraction);
    }
  }
}

async function triggerSubsyncFromConfig(): Promise<void> {
  if (subsyncInProgress) {
    showMpvOsd("Subsync already running");
    return;
  }
  const resolved = getSubsyncConfig(getResolvedConfig().subsync);
  try {
    if (resolved.defaultMode === "manual") {
      await openSubsyncManualPicker();
      showMpvOsd("Subsync: choose engine and source");
      return;
    }

    subsyncInProgress = true;
    const result = await runWithSubsyncSpinner(() => runSubsyncAuto());
    showMpvOsd(result.message);
  } catch (error) {
    showMpvOsd(`Subsync failed: ${(error as Error).message}`);
  } finally {
    subsyncInProgress = false;
  }
}

function formatLangScore(name: string, pref: JimakuLanguagePreference): number {
  if (pref === "none") return 0;
  const upper = name.toUpperCase();
  const hasJa =
    /(^|[\W_])JA([\W_]|$)/.test(upper) ||
    /(^|[\W_])JPN([\W_]|$)/.test(upper) ||
    upper.includes(".JA.");
  const hasEn =
    /(^|[\W_])EN([\W_]|$)/.test(upper) ||
    /(^|[\W_])ENG([\W_]|$)/.test(upper) ||
    upper.includes(".EN.");
  if (pref === "ja") {
    if (hasJa) return 2;
    if (hasEn) return 1;
  } else if (pref === "en") {
    if (hasEn) return 2;
    if (hasJa) return 1;
  }
  return 0;
}

function sortJimakuFiles(
  files: JimakuFileEntry[],
  pref: JimakuLanguagePreference,
): JimakuFileEntry[] {
  if (pref === "none") return files;
  return [...files].sort((a, b) => {
    const scoreDiff =
      formatLangScore(b.name, pref) - formatLangScore(a.name, pref);
    if (scoreDiff !== 0) return scoreDiff;
    return a.name.localeCompare(b.name);
  });
}

function isRemoteMediaPath(mediaPath: string): boolean {
  return /^[a-z][a-z0-9+.-]*:\/\//i.test(mediaPath);
}

async function downloadToFile(
  url: string,
  destPath: string,
  headers: Record<string, string>,
  redirectCount = 0,
): Promise<JimakuDownloadResult> {
  if (redirectCount > 3) {
    return {
      ok: false,
      error: { error: "Too many redirects while downloading subtitle." },
    };
  }

  return new Promise((resolve) => {
    const parsedUrl = new URL(url);
    const transport = parsedUrl.protocol === "https:" ? https : http;

    const req = transport.get(parsedUrl, { headers }, (res) => {
      const status = res.statusCode || 0;
      if ([301, 302, 303, 307, 308].includes(status) && res.headers.location) {
        const redirectUrl = new URL(res.headers.location, parsedUrl).toString();
        res.resume();
        downloadToFile(redirectUrl, destPath, headers, redirectCount + 1).then(
          resolve,
        );
        return;
      }

      if (status < 200 || status >= 300) {
        res.resume();
        resolve({
          ok: false,
          error: {
            error: `Failed to download subtitle (HTTP ${status}).`,
            code: status,
          },
        });
        return;
      }

      const fileStream = fs.createWriteStream(destPath);
      res.pipe(fileStream);
      fileStream.on("finish", () => {
        fileStream.close(() => {
          resolve({ ok: true, path: destPath });
        });
      });
      fileStream.on("error", (err) => {
        resolve({
          ok: false,
          error: {
            error: `Failed to save subtitle: ${(err as Error).message}`,
          },
        });
      });
    });

    req.on("error", (err) => {
      resolve({
        ok: false,
        error: { error: `Download request failed: ${(err as Error).message}` },
      });
    });
  });
}

function cancelPendingMultiCopy(): void {
  if (!pendingMultiCopy) return;

  pendingMultiCopy = false;
  if (pendingMultiCopyTimeout) {
    clearTimeout(pendingMultiCopyTimeout);
    pendingMultiCopyTimeout = null;
  }

  // Unregister digit and escape shortcuts
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

  // Register digit shortcuts 1-9
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

  // Register Escape to cancel
  if (
    globalShortcut.register("Escape", () => {
      cancelPendingMultiCopy();
      showMpvOsd("Cancelled");
    })
  ) {
    multiCopyEscapeShortcut = "Escape";
  }

  // Set timeout
  pendingMultiCopyTimeout = setTimeout(() => {
    cancelPendingMultiCopy();
    showMpvOsd("Copy timeout");
  }, timeoutMs);

  showMpvOsd("Copy how many lines? Press 1-9 (Esc to cancel)");
}

function handleMultiCopyDigit(count: number): void {
  if (!pendingMultiCopy || !subtitleTimingTracker) return;

  cancelPendingMultiCopy();

  // Check if we have enough history
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
  let registeredAny = false;
  const registerOverlayShortcut = (
    accelerator: string,
    handler: () => void,
    label: string,
  ): void => {
    if (isGlobalShortcutRegisteredSafe(accelerator)) {
      registeredAny = true;
      return;
    }
    const ok = globalShortcut.register(accelerator, handler);
    if (!ok) {
      console.warn(
        `Failed to register overlay shortcut ${label}: ${accelerator}`,
      );
      return;
    }
    registeredAny = true;
  };

  if (shortcuts.copySubtitleMultiple) {
    registerOverlayShortcut(
      shortcuts.copySubtitleMultiple,
      () => {
        startPendingMultiCopy(shortcuts.multiCopyTimeoutMs);
      },
      "copySubtitleMultiple",
    );
  }

  if (shortcuts.copySubtitle) {
    registerOverlayShortcut(
      shortcuts.copySubtitle,
      () => {
        copyCurrentSubtitle();
      },
      "copySubtitle",
    );
  }

  if (shortcuts.triggerFieldGrouping) {
    registerOverlayShortcut(
      shortcuts.triggerFieldGrouping,
      () => {
        triggerFieldGrouping().catch((err) => {
          console.error("triggerFieldGrouping failed:", err);
          showMpvOsd(`Field grouping failed: ${(err as Error).message}`);
        });
      },
      "triggerFieldGrouping",
    );
  }

  if (shortcuts.triggerSubsync) {
    registerOverlayShortcut(
      shortcuts.triggerSubsync,
      () => {
        triggerSubsyncFromConfig().catch((err) => {
          console.error("triggerSubsyncFromConfig failed:", err);
          showMpvOsd(`Subsync failed: ${(err as Error).message}`);
        });
      },
      "triggerSubsync",
    );
  }

  if (shortcuts.mineSentence) {
    registerOverlayShortcut(
      shortcuts.mineSentence,
      () => {
        mineSentenceCard().catch((err) => {
          console.error("mineSentenceCard failed:", err);
          showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
        });
      },
      "mineSentence",
    );
  }

  if (shortcuts.mineSentenceMultiple) {
    registerOverlayShortcut(
      shortcuts.mineSentenceMultiple,
      () => {
        startPendingMineSentenceMultiple(shortcuts.multiCopyTimeoutMs);
      },
      "mineSentenceMultiple",
    );
  }

  if (shortcuts.toggleSecondarySub) {
    registerOverlayShortcut(
      shortcuts.toggleSecondarySub,
      () => cycleSecondarySubMode(),
      "toggleSecondarySub",
    );
  }

  if (shortcuts.updateLastCardFromClipboard) {
    registerOverlayShortcut(
      shortcuts.updateLastCardFromClipboard,
      () => {
        updateLastCardFromClipboard().catch((err) => {
          console.error("updateLastCardFromClipboard failed:", err);
          showMpvOsd(`Update failed: ${(err as Error).message}`);
        });
      },
      "updateLastCardFromClipboard",
    );
  }

  if (shortcuts.markAudioCard) {
    registerOverlayShortcut(
      shortcuts.markAudioCard,
      () => {
        markLastCardAsAudioCard().catch((err) => {
          console.error("markLastCardAsAudioCard failed:", err);
          showMpvOsd(`Audio card failed: ${(err as Error).message}`);
        });
      },
      "markAudioCard",
    );
  }

  if (shortcuts.openRuntimeOptions) {
    registerOverlayShortcut(
      shortcuts.openRuntimeOptions,
      () => {
        openRuntimeOptionsPalette();
      },
      "openRuntimeOptions",
    );
  }

  shortcutsRegistered = registeredAny;
}

function unregisterOverlayShortcuts(): void {
  if (!shortcutsRegistered) return;

  cancelPendingMultiCopy();
  cancelPendingMineSentenceMultiple();

  const shortcuts = getConfiguredShortcuts();

  if (shortcuts.copySubtitle) {
    globalShortcut.unregister(shortcuts.copySubtitle);
  }
  if (shortcuts.copySubtitleMultiple) {
    globalShortcut.unregister(shortcuts.copySubtitleMultiple);
  }
  if (shortcuts.updateLastCardFromClipboard) {
    globalShortcut.unregister(shortcuts.updateLastCardFromClipboard);
  }
  if (shortcuts.triggerFieldGrouping) {
    globalShortcut.unregister(shortcuts.triggerFieldGrouping);
  }
  if (shortcuts.triggerSubsync) {
    globalShortcut.unregister(shortcuts.triggerSubsync);
  }
  if (shortcuts.mineSentence) {
    globalShortcut.unregister(shortcuts.mineSentence);
  }
  if (shortcuts.mineSentenceMultiple) {
    globalShortcut.unregister(shortcuts.mineSentenceMultiple);
  }
  if (shortcuts.toggleSecondarySub) {
    globalShortcut.unregister(shortcuts.toggleSecondarySub);
  }
  if (shortcuts.markAudioCard) {
    globalShortcut.unregister(shortcuts.markAudioCard);
  }
  if (shortcuts.openRuntimeOptions) {
    globalShortcut.unregister(shortcuts.openRuntimeOptions);
  }

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
  console.log(
    "updateVisibleOverlayVisibility called, visibleOverlayVisible:",
    visibleOverlayVisible,
  );
  if (!mainWindow || mainWindow.isDestroyed()) {
    console.log("mainWindow not available");
    return;
  }

  if (!visibleOverlayVisible) {
    console.log("Hiding visible overlay");
    mainWindow.hide();

    if (
      shouldBindVisibleOverlayToMpvSubVisibility() &&
      previousSecondarySubVisibility !== null &&
      mpvClient &&
      mpvClient.connected
    ) {
      mpvClient.send({
        command: [
          "set_property",
          "secondary-sub-visibility",
          previousSecondarySubVisibility ? "yes" : "no",
        ],
      });
      previousSecondarySubVisibility = null;
    } else if (!shouldBindVisibleOverlayToMpvSubVisibility()) {
      previousSecondarySubVisibility = null;
    }
    syncOverlayShortcuts();
  } else {
    console.log(
      "Should show visible overlay, isTracking:",
      windowTracker?.isTracking(),
    );

    if (
      shouldBindVisibleOverlayToMpvSubVisibility() &&
      mpvClient &&
      mpvClient.connected
    ) {
      mpvClient.send({
        command: ["get_property", "secondary-sub-visibility"],
        request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
      });
    }

    if (windowTracker && windowTracker.isTracking()) {
      trackerNotReadyWarningShown = false;
      const geometry = windowTracker.getGeometry();
      console.log("Geometry:", geometry);
      if (geometry) {
        updateOverlayBounds(geometry);
      }
      console.log("Showing visible overlay mainWindow");
      ensureOverlayWindowLevel(mainWindow);
      mainWindow.show();
      mainWindow.focus();
      enforceOverlayLayerOrder();
      syncOverlayShortcuts();
    } else if (!windowTracker) {
      trackerNotReadyWarningShown = false;
      ensureOverlayWindowLevel(mainWindow);
      mainWindow.show();
      mainWindow.focus();
      enforceOverlayLayerOrder();
      syncOverlayShortcuts();
    } else {
      if (!trackerNotReadyWarningShown) {
        console.warn(
          "Window tracker exists but is not tracking yet; using fallback bounds until tracking starts",
        );
        trackerNotReadyWarningShown = true;
      }
      const cursorPoint = screen.getCursorScreenPoint();
      const display = screen.getDisplayNearestPoint(cursorPoint);
      const fallbackBounds = display.workArea;
      updateOverlayBounds({
        x: fallbackBounds.x,
        y: fallbackBounds.y,
        width: fallbackBounds.width,
        height: fallbackBounds.height,
      });
      ensureOverlayWindowLevel(mainWindow);
      mainWindow.show();
      mainWindow.focus();
      enforceOverlayLayerOrder();
      syncOverlayShortcuts();
    }
  }
}

function updateInvisibleOverlayVisibility(): void {
  if (!invisibleWindow || invisibleWindow.isDestroyed()) {
    return;
  }

  // When the visible overlay is shown, keep the invisible layer hidden to
  // avoid it intercepting or rerouting pointer interactions.
  if (visibleOverlayVisible) {
    invisibleWindow.hide();
    syncOverlayShortcuts();
    return;
  }

  const showInvisibleWithoutFocus = (): void => {
    ensureOverlayWindowLevel(invisibleWindow!);
    if (typeof invisibleWindow!.showInactive === "function") {
      invisibleWindow!.showInactive();
    } else {
      invisibleWindow!.show();
    }
    enforceOverlayLayerOrder();
  };

  if (!invisibleOverlayVisible) {
    invisibleWindow.hide();
    syncOverlayShortcuts();
    return;
  }

  if (windowTracker && windowTracker.isTracking()) {
    const geometry = windowTracker.getGeometry();
    if (geometry) {
      updateOverlayBounds(geometry);
    }
    showInvisibleWithoutFocus();
    syncOverlayShortcuts();
    return;
  }

  if (!windowTracker) {
    showInvisibleWithoutFocus();
    syncOverlayShortcuts();
    return;
  }

  const cursorPoint = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(cursorPoint);
  const fallbackBounds = display.workArea;
  updateOverlayBounds({
    x: fallbackBounds.x,
    y: fallbackBounds.y,
    width: fallbackBounds.width,
    height: fallbackBounds.height,
  });
  showInvisibleWithoutFocus();
  syncOverlayShortcuts();
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

ipcMain.on(
  "set-ignore-mouse-events",
  (
    event: IpcMainEvent,
    ignore: boolean,
    options: { forward?: boolean } = {},
  ) => {
    const senderWindow = BrowserWindow.fromWebContents(event.sender);
    if (senderWindow && !senderWindow.isDestroyed()) {
      if (
        senderWindow === invisibleWindow &&
        visibleOverlayVisible &&
        !invisibleWindow.isDestroyed()
      ) {
        invisibleWindow.setIgnoreMouseEvents(true, { forward: true });
      } else {
        senderWindow.setIgnoreMouseEvents(ignore, options);
      }
    }
  },
);

ipcMain.on(
  "overlay:modal-closed",
  (_event: IpcMainEvent, modal: OverlayHostedModal) => {
    if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
    restoreVisibleOverlayOnModalClose.delete(modal);
    if (restoreVisibleOverlayOnModalClose.size === 0) {
      setVisibleOverlayVisible(false);
    }
  },
);

ipcMain.on("open-yomitan-settings", () => {
  openYomitanSettings();
});

ipcMain.on("quit-app", () => {
  app.quit();
});

ipcMain.on("toggle-dev-tools", () => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.toggleDevTools();
  }
});

ipcMain.handle("get-overlay-visibility", () => {
  return visibleOverlayVisible;
});

ipcMain.on("toggle-overlay", () => {
  toggleVisibleOverlay();
});

ipcMain.handle("get-visible-overlay-visibility", () => {
  return visibleOverlayVisible;
});

ipcMain.handle("get-invisible-overlay-visibility", () => {
  return invisibleOverlayVisible;
});

ipcMain.handle("get-current-subtitle", async () => {
  return await tokenizeSubtitle(currentSubText);
});

ipcMain.handle("get-current-subtitle-ass", () => {
  return currentSubAssText;
});

ipcMain.handle("get-mpv-subtitle-render-metrics", () => {
  return mpvSubtitleRenderMetrics;
});

ipcMain.handle("get-subtitle-position", () => {
  return loadSubtitlePosition();
});

ipcMain.handle("get-subtitle-style", () => {
  const config = getResolvedConfig();
  return config.subtitleStyle ?? null;
});

ipcMain.on(
  "save-subtitle-position",
  (_event: IpcMainEvent, position: SubtitlePosition) => {
    saveSubtitlePosition(position);
  },
);

ipcMain.handle("get-mecab-status", () => {
  if (mecabTokenizer) {
    return mecabTokenizer.getStatus();
  }
  return { available: false, enabled: false, path: null };
});

ipcMain.on("set-mecab-enabled", (_event: IpcMainEvent, enabled: boolean) => {
  if (mecabTokenizer) {
    mecabTokenizer.setEnabled(enabled);
  }
});

ipcMain.on(
  "mpv-command",
  (_event: IpcMainEvent, command: (string | number)[]) => {
    const first = typeof command[0] === "string" ? command[0] : "";
    if (first === SPECIAL_COMMANDS.SUBSYNC_TRIGGER) {
      triggerSubsyncFromConfig();
      return;
    }

    if (first === SPECIAL_COMMANDS.RUNTIME_OPTIONS_OPEN) {
      openRuntimeOptionsPalette();
      return;
    }

    if (first.startsWith(SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX)) {
      if (!runtimeOptionsManager) return;
      const [, idToken, directionToken] = first.split(":");
      const id = idToken as RuntimeOptionId;
      const direction: 1 | -1 = directionToken === "prev" ? -1 : 1;
      const result = applyRuntimeOptionResult(
        runtimeOptionsManager.cycleOption(id, direction),
      );
      if (!result.ok && result.error) {
        showMpvOsd(result.error);
      }
      return;
    }

    if (mpvClient && mpvClient.connected) {
      if (first === SPECIAL_COMMANDS.REPLAY_SUBTITLE) {
        mpvClient.replayCurrentSubtitle();
      } else if (first === SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE) {
        mpvClient.playNextSubtitle();
      } else {
        mpvClient.send({ command });
      }
    }
  },
);

ipcMain.handle("get-keybindings", () => {
  return keybindings;
});

ipcMain.handle("get-secondary-sub-mode", () => {
  return secondarySubMode;
});

ipcMain.handle("get-current-secondary-sub", () => {
  return mpvClient?.currentSecondarySubText || "";
});

ipcMain.handle(
  "subsync:run-manual",
  async (_event, request: SubsyncManualRunRequest): Promise<SubsyncResult> => {
    if (subsyncInProgress) {
      const busy = "Subsync already running";
      showMpvOsd(busy);
      return { ok: false, message: busy };
    }
    try {
      subsyncInProgress = true;
      const result = await runWithSubsyncSpinner(() =>
        runSubsyncManual(request),
      );
      showMpvOsd(result.message);
      return result;
    } catch (error) {
      const message = `Subsync failed: ${(error as Error).message}`;
      showMpvOsd(message);
      return { ok: false, message };
    } finally {
      subsyncInProgress = false;
    }
  },
);

ipcMain.handle("get-anki-connect-status", () => {
  return ankiIntegration !== null;
});

ipcMain.handle("runtime-options:get", (): RuntimeOptionState[] => {
  return getRuntimeOptionsState();
});

ipcMain.handle(
  "runtime-options:set",
  (
    _event,
    id: RuntimeOptionId,
    value: RuntimeOptionValue,
  ): RuntimeOptionApplyResult => {
    if (!runtimeOptionsManager) {
      return { ok: false, error: "Runtime options manager unavailable" };
    }
    const result = applyRuntimeOptionResult(
      runtimeOptionsManager.setOptionValue(id, value),
    );
    if (!result.ok && result.error) {
      showMpvOsd(result.error);
    }
    return result;
  },
);

ipcMain.handle(
  "runtime-options:cycle",
  (
    _event,
    id: RuntimeOptionId,
    direction: 1 | -1,
  ): RuntimeOptionApplyResult => {
    if (!runtimeOptionsManager) {
      return { ok: false, error: "Runtime options manager unavailable" };
    }
    const result = applyRuntimeOptionResult(
      runtimeOptionsManager.cycleOption(id, direction),
    );
    if (!result.ok && result.error) {
      showMpvOsd(result.error);
    }
    return result;
  },
);

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

        // Treat the visible overlay as a temporary modal host and then restore
        // the user's previous overlay visibility state.
        if (!previousVisibleOverlay && visibleOverlayVisible) {
          setVisibleOverlayVisible(false);
        }
        if (invisibleOverlayVisible !== previousInvisibleOverlay) {
          setInvisibleOverlayVisible(previousInvisibleOverlay);
        }
      };

      fieldGroupingResolver = finish;
      // Manual Kiku flow is rendered only in the visible overlay window.
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
  if (!mainWindow || mainWindow.isDestroyed()) return false;
  const wasVisible = visibleOverlayVisible;
  if (!visibleOverlayVisible) {
    setVisibleOverlayVisible(true);
  }
  if (!wasVisible && options?.restoreOnModalClose) {
    restoreVisibleOverlayOnModalClose.add(options.restoreOnModalClose);
  }
  if (payload === undefined) {
    mainWindow.webContents.send(channel);
  } else {
    mainWindow.webContents.send(channel, payload);
  }
  return true;
}

function showDesktopNotification(
  title: string,
  options: { body?: string; icon?: string },
): void {
  const notificationOptions: {
    title: string;
    body?: string;
    icon?: Electron.NativeImage | string;
  } = { title };

  if (options.body) {
    notificationOptions.body = options.body;
  }

  if (options.icon) {
    // Check if it's a file path (starts with / on Linux/Mac, or drive letter on Windows)
    const isFilePath =
      typeof options.icon === "string" &&
      (options.icon.startsWith("/") || /^[a-zA-Z]:[\\/]/.test(options.icon));

    if (isFilePath) {
      // File path - preferred for Linux/Wayland compatibility
      // Verify file exists before using
      if (fs.existsSync(options.icon)) {
        notificationOptions.icon = options.icon;
      } else {
        console.warn("Notification icon file not found:", options.icon);
      }
    } else if (
      typeof options.icon === "string" &&
      options.icon.startsWith("data:image/")
    ) {
      // Data URL fallback - decode to nativeImage
      const base64Data = options.icon.replace(/^data:image\/\w+;base64,/, "");
      try {
        const image = nativeImage.createFromBuffer(
          Buffer.from(base64Data, "base64"),
        );
        if (image.isEmpty()) {
          console.warn(
            "Notification icon created from base64 is empty - image format may not be supported by Electron",
          );
        } else {
          notificationOptions.icon = image;
        }
      } catch (err) {
        console.error("Failed to create notification icon from base64:", err);
      }
    } else {
      // Unknown format, try to use as-is
      notificationOptions.icon = options.icon;
    }
  }

  const notification = new Notification(notificationOptions);
  notification.show();
}

ipcMain.on(
  "set-anki-connect-enabled",
  (_event: IpcMainEvent, enabled: boolean) => {
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
);

ipcMain.on("clear-anki-connect-history", () => {
  if (subtitleTimingTracker) {
    subtitleTimingTracker.cleanup();
    console.log("AnkiConnect subtitle timing history cleared");
  }
});

ipcMain.on(
  "kiku:field-grouping-respond",
  (_event: IpcMainEvent, choice: KikuFieldGroupingChoice) => {
    if (fieldGroupingResolver) {
      fieldGroupingResolver(choice);
      fieldGroupingResolver = null;
    }
  },
);

ipcMain.handle(
  "kiku:build-merge-preview",
  async (
    _event,
    request: KikuMergePreviewRequest,
  ): Promise<KikuMergePreviewResponse> => {
    if (!ankiIntegration) {
      return { ok: false, error: "AnkiConnect integration not enabled" };
    }
    return ankiIntegration.buildFieldGroupingPreview(
      request.keepNoteId,
      request.deleteNoteId,
      request.deleteDuplicate,
    );
  },
);

ipcMain.handle("jimaku:get-media-info", (): JimakuMediaInfo => {
  return parseMediaInfo(currentMediaPath);
});

ipcMain.handle(
  "jimaku:search-entries",
  async (
    _event,
    query: JimakuSearchQuery,
  ): Promise<JimakuApiResponse<JimakuEntry[]>> => {
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
);

ipcMain.handle(
  "jimaku:list-files",
  async (
    _event,
    query: JimakuFilesQuery,
  ): Promise<JimakuApiResponse<JimakuFileEntry[]>> => {
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
);

ipcMain.handle(
  "jimaku:download-file",
  async (_event, query: JimakuDownloadQuery): Promise<JimakuDownloadResult> => {
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

    if (!currentMediaPath) {
      return { ok: false, error: { error: "No media file loaded in MPV." } };
    }

    if (isRemoteMediaPath(currentMediaPath)) {
      return {
        ok: false,
        error: { error: "Cannot download subtitles for remote media paths." },
      };
    }

    const mediaDir = path.dirname(path.resolve(currentMediaPath));
    const safeName = path.basename(query.name);
    if (!safeName) {
      return { ok: false, error: { error: "Invalid subtitle filename." } };
    }

    const ext = path.extname(safeName);
    const baseName = ext ? safeName.slice(0, -ext.length) : safeName;
    let targetPath = path.join(mediaDir, safeName);
    if (fs.existsSync(targetPath)) {
      targetPath = path.join(
        mediaDir,
        `${baseName} (jimaku-${query.entryId})${ext}`,
      );
      let counter = 2;
      while (fs.existsSync(targetPath)) {
        targetPath = path.join(
          mediaDir,
          `${baseName} (jimaku-${query.entryId}-${counter})${ext}`,
        );
        counter += 1;
      }
    }

    console.log(
      `[jimaku] download-file name="${query.name}" entryId=${query.entryId}`,
    );
    const result = await downloadToFile(query.url, targetPath, {
      Authorization: apiKey,
      "User-Agent": "SubMiner",
    });

    if (result.ok) {
      console.log(`[jimaku] download-file saved to ${result.path}`);
      if (mpvClient && mpvClient.connected) {
        mpvClient.send({ command: ["sub-add", result.path, "select"] });
      }
    } else {
      console.error(
        `[jimaku] download-file failed: ${result.error?.error ?? "unknown error"}`,
      );
    }

    return result;
  },
);
