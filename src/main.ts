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
import * as childProcess from "child_process";
import WebSocket from "ws";
import { parse as parseJsonc } from "jsonc-parser";
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
  KikuFieldGroupingRequestData,
  KikuFieldGroupingChoice,
  KikuMergePreviewRequest,
  KikuMergePreviewResponse,
} from "./types";
import { SubtitleTimingTracker } from "./subtitle-timing-tracker";
import { AnkiIntegration } from "./anki-integration";

const DEFAULT_TEXTHOOKER_PORT = 5174;
const DEFAULT_WEBSOCKET_PORT = 6677;
const DEFAULT_SUBTITLE_FONT_SIZE = 24;
let texthookerServer: http.Server | null = null;
let subtitleWebSocketServer: WebSocket.Server | null = null;
const CONFIG_DIR = path.join(os.homedir(), ".config", "SubMiner");
const USER_DATA_PATH = CONFIG_DIR;
const isDev = process.argv.includes("--dev");

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
  settings: boolean;
  show: boolean;
  hide: boolean;
  texthooker: boolean;
  help: boolean;
  autoStartOverlay: boolean;
  socketPath?: string;
  backend?: string;
  texthookerPort?: number;
}

function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = {
    start: false,
    stop: false,
    toggle: false,
    settings: false,
    show: false,
    hide: false,
    texthooker: false,
    help: false,
    autoStartOverlay: false,
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
    else if (arg === "--settings" || arg === "--yomitan") args.settings = true;
    else if (arg === "--show") args.show = true;
    else if (arg === "--hide") args.hide = true;
    else if (arg === "--texthooker") args.texthooker = true;
    else if (arg === "--auto-start-overlay") args.autoStartOverlay = true;
    else if (arg === "--help") args.help = true;
    else if (arg.startsWith("--socket=")) {
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
  --start               Start the overlay app (default on GUI launch)
  --stop                Stop the running overlay app
  --toggle              Toggle subtitle overlay visibility
  --settings            Open Yomitan settings window
  --texthooker          Start texthooker server and open browser
  --show                Force show overlay
  --hide                Force hide overlay
  --auto-start-overlay  Auto-hide mpv subtitles on connect (show overlay)
  --socket PATH         Override MPV IPC socket/pipe path
  --backend BACKEND     Override window tracker backend (auto, hyprland, sway, x11, macos)
  --port PORT           Texthooker server port (default: ${DEFAULT_TEXTHOOKER_PORT})
  --dev                 Run in development mode
  --help                Show this help
`);
}

function hasExplicitCommand(args: CliArgs): boolean {
  return (
    args.start ||
    args.stop ||
    args.toggle ||
    args.settings ||
    args.show ||
    args.hide ||
    args.texthooker ||
    args.help
  );
}

function shouldStartApp(args: CliArgs): boolean {
  if (args.stop && !args.start) return false;
  if (
    args.start ||
    args.toggle ||
    args.settings ||
    args.show ||
    args.hide ||
    args.texthooker
  ) {
    return true;
  }
  return !hasExplicitCommand(args) && !args.help;
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
let yomitanExt: Extension | null = null;
let yomitanSettingsWindow: BrowserWindow | null = null;
let mpvClient: MpvIpcClient | null = null;
let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
let currentSubText = "";
let overlayVisible = false;
let windowTracker: BaseWindowTracker | null = null;
let subtitlePosition: SubtitlePosition | null = null;
let currentMediaPath: string | null = null;
let mecabTokenizer: MecabTokenizer | null = null;
let keybindings: Keybinding[] = [];
let subtitleTimingTracker: SubtitleTimingTracker | null = null;
let ankiIntegration: AnkiIntegration | null = null;
let secondarySubMode: SecondarySubMode = "hover";
let previousSecondarySubVisibility: boolean | null = null;

// Shortcut state tracking
let shortcutsRegistered = false;
let pendingMultiCopy = false;
let pendingMultiCopyTimeout: ReturnType<typeof setTimeout> | null = null;
let multiCopyDigitShortcuts: string[] = [];
let multiCopyEscapeShortcut: string | null = null;
let pendingMineSentenceMultiple = false;
let pendingMineSentenceMultipleTimeout: ReturnType<typeof setTimeout> | null =
  null;
let mineSentenceDigitShortcuts: string[] = [];
let mineSentenceEscapeShortcut: string | null = null;
let fieldGroupingResolver: ((choice: KikuFieldGroupingChoice) => void) | null =
  null;

const DEFAULT_KEYBINDINGS: Keybinding[] = [
  { key: "Space", command: ["cycle", "pause"] },
  { key: "Shift+KeyH", command: ["sub-seek", -1] },
  { key: "Shift+KeyL", command: ["sub-seek", 1] },
  { key: "Ctrl+Shift+KeyH", command: ["__replay-subtitle"] },
  { key: "Ctrl+Shift+KeyL", command: ["__play-next-subtitle"] },
];

const CONFIG_FILE_JSONC = path.join(CONFIG_DIR, "config.jsonc");
const CONFIG_FILE_JSON = path.join(CONFIG_DIR, "config.json");
const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, "subtitle-positions");
const DEFAULT_JIMAKU_BASE_URL = "https://jimaku.cc";
const DEFAULT_JIMAKU_MAX_RESULTS = 10;
const DEFAULT_JIMAKU_LANGUAGE_PREF: JimakuLanguagePreference = "ja";

function getConfigFilePath(): string {
  if (fs.existsSync(CONFIG_FILE_JSONC)) return CONFIG_FILE_JSONC;
  if (fs.existsSync(CONFIG_FILE_JSON)) return CONFIG_FILE_JSON;
  return CONFIG_FILE_JSONC;
}

interface LoadConfigResult {
  success: boolean;
  config: Config;
}

function loadConfig(): LoadConfigResult {
  try {
    const configPath = getConfigFilePath();
    if (fs.existsSync(configPath)) {
      const data = fs.readFileSync(configPath, "utf-8");
      const config = configPath.endsWith(".jsonc")
        ? parseJsonc(data)
        : JSON.parse(data);
      return { success: true, config: config as Config };
    }
    return { success: true, config: {} };
  } catch (err) {
    console.error("Failed to load config:", (err as Error).message);
    return { success: false, config: {} };
  }
}

function saveConfig(config: Config): void {
  try {
    if (!fs.existsSync(CONFIG_DIR)) {
      fs.mkdirSync(CONFIG_DIR, { recursive: true });
    }
    const configPath = getConfigFilePath();
    fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
  } catch (err) {
    console.error("Failed to save config:", (err as Error).message);
  }
}

function getJimakuConfig(): JimakuConfig {
  const { config } = loadConfig();
  return config.jimaku ?? {};
}

function getJimakuBaseUrl(): string {
  const config = getJimakuConfig();
  return config.apiBaseUrl || DEFAULT_JIMAKU_BASE_URL;
}

function getJimakuLanguagePreference(): JimakuLanguagePreference {
  const config = getJimakuConfig();
  return config.languagePreference || DEFAULT_JIMAKU_LANGUAGE_PREF;
}

function getJimakuMaxEntryResults(): number {
  const config = getJimakuConfig();
  const value = config.maxEntryResults;
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  return DEFAULT_JIMAKU_MAX_RESULTS;
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
  const hash = crypto.createHash("sha256").update(mediaPath).digest("hex");
  return path.join(SUBTITLE_POSITIONS_DIR, `${hash}.json`);
}

function loadSubtitlePosition(): SubtitlePosition | null {
  if (!currentMediaPath) {
    subtitlePosition = null;
    return subtitlePosition;
  }

  try {
    const positionPath = getSubtitlePositionFilePath(currentMediaPath);
    if (!fs.existsSync(positionPath)) {
      subtitlePosition = null;
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
      subtitlePosition = null;
    }
  } catch (err) {
    console.error("Failed to load subtitle position:", (err as Error).message);
    subtitlePosition = null;
  }

  return subtitlePosition;
}

function saveSubtitlePosition(position: SubtitlePosition): void {
  subtitlePosition = position;
  if (!currentMediaPath) {
    console.error("Refusing to save subtitle position - no media path yet");
    return;
  }

  try {
    if (!fs.existsSync(SUBTITLE_POSITIONS_DIR)) {
      fs.mkdirSync(SUBTITLE_POSITIONS_DIR, { recursive: true });
    }
    const positionPath = getSubtitlePositionFilePath(currentMediaPath);
    fs.writeFileSync(positionPath, JSON.stringify(position, null, 2));
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
  const position = loadSubtitlePosition();
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send("subtitle-position:set", position);
  }
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

function loadKeybindings(): Keybinding[] {
  const { config } = loadConfig();
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
let mpvSocketPath = initialArgs.socketPath ?? getDefaultSocketPath();
let texthookerPort = initialArgs.texthookerPort ?? DEFAULT_TEXTHOOKER_PORT;
const backendOverride = initialArgs.backend ?? null;
const autoStartOverlay = initialArgs.autoStartOverlay;

const gotTheLock = app.requestSingleInstanceLock();

if (!gotTheLock) {
  app.quit();
} else {
  app.on("second-instance", (_event, argv) => {
    handleCliCommand(parseArgs(argv));
  });
  if (initialArgs.help && !shouldStartApp(initialArgs)) {
    printHelp();
    app.quit();
  } else if (!shouldStartApp(initialArgs)) {
    console.error("No running instance. Use --start to launch the app.");
    app.quit();
  } else {
    app.whenReady().then(async () => {
      loadSubtitlePosition();
      loadKeybindings();

      mpvClient = new MpvIpcClient(mpvSocketPath);
      mpvClient.connect();

      const { config } = loadConfig();
      secondarySubMode = config.secondarySub?.defaultMode ?? "hover";
      const wsConfig = config.websocket || {};
      const wsEnabled = wsConfig.enabled ?? "auto";
      const wsPort = wsConfig.port || DEFAULT_WEBSOCKET_PORT;

      if (wsEnabled === true || (wsEnabled === "auto" && !hasMpvWebsocket())) {
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
      createMainWindow();
      registerGlobalShortcuts();

      windowTracker = createWindowTracker(backendOverride);
      if (windowTracker) {
        windowTracker.onGeometryChange = (geometry: WindowGeometry) => {
          updateOverlayBounds(geometry);
        };
        windowTracker.onWindowFound = (geometry: WindowGeometry) => {
          console.log("MPV window found:", geometry);
          updateOverlayBounds(geometry);
          if (overlayVisible) {
            updateOverlayVisibility();
          }
        };
        windowTracker.onWindowLost = () => {
          console.log("MPV window lost");
          if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.hide();
          }
          unregisterOverlayShortcuts();
        };
        windowTracker.start();
      }

      if (config.ankiConnect?.enabled && subtitleTimingTracker && mpvClient) {
        ankiIntegration = new AnkiIntegration(
          config.ankiConnect,
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
      if (BrowserWindow.getAllWindows().length === 0) {
        createMainWindow();
      }
    });
  }
}

function handleCliCommand(args: CliArgs): void {
  if (args.socketPath !== undefined) {
    if (mpvClient) {
      console.warn(
        "Ignoring --socket override because the IPC client is already running.",
      );
    } else {
      mpvSocketPath = args.socketPath;
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
  } else if (args.toggle) {
    toggleOverlay();
  } else if (args.settings) {
    setTimeout(() => {
      openYomitanSettings();
    }, 1000);
  } else if (args.show) {
    setOverlayVisible(true);
  } else if (args.hide) {
    setOverlayVisible(false);
  } else if (args.texthooker) {
    if (!texthookerServer) {
      startTexthookerServer(texthookerPort);
    }
    const { config } = loadConfig();
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
  handleCliCommand(initialArgs);
}

interface MpvMessage {
  event?: string;
  name?: string;
  data?: unknown;
  request_id?: number;
}

const MPV_REQUEST_ID_SUBTEXT = 101;
const MPV_REQUEST_ID_PATH = 102;
const MPV_REQUEST_ID_SECONDARY_SUBTEXT = 103;
const MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY = 104;
const MPV_REQUEST_ID_AID = 105;
const MPV_REQUEST_ID_TRACK_LIST_SECONDARY = 200;
const MPV_REQUEST_ID_TRACK_LIST_AUDIO = 201;

class MpvIpcClient implements MpvClient {
  private socketPath: string;
  public socket: net.Socket | null = null;
  private buffer = "";
  public connected = false;
  private reconnectAttempt = 0;
  private firstConnection = true;
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

  constructor(socketPath: string) {
    this.socketPath = socketPath;
  }

  connect(): void {
    if (this.socket) {
      this.socket.destroy();
    }

    this.socket = new net.Socket();

    this.socket.on("connect", () => {
      console.log("Connected to MPV socket");
      this.connected = true;
      this.reconnectAttempt = 0;
      this.subscribeToProperties();
      this.getInitialState();

      const shouldAutoStart =
        autoStartOverlay || loadConfig().config.auto_start_overlay === true;
      if (this.firstConnection && shouldAutoStart) {
        console.log("Auto-starting overlay, hiding mpv subtitles");
        setTimeout(() => {
          setOverlayVisible(true);
        }, 100);
      } else {
        this.setSubVisibility(!overlayVisible);
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
    });

    this.socket.on("close", () => {
      console.log("MPV socket closed");
      this.connected = false;
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
    if (attempt < 2) {
      delay = 200;
    } else if (attempt < 4) {
      delay = 500;
    } else if (attempt < 6) {
      delay = 1000;
    } else {
      delay = 2000;
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
        if (mainWindow && !mainWindow.isDestroyed()) {
          const subtitleData = await tokenizeSubtitle(currentSubText);
          mainWindow.webContents.send("subtitle:set", subtitleData);
        }
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
        if (mainWindow && !mainWindow.isDestroyed()) {
          mainWindow.webContents.send(
            "secondary-subtitle:set",
            this.currentSecondarySubText,
          );
        }
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
      }
    } else if (msg.data !== undefined && msg.request_id) {
      if (msg.request_id === MPV_REQUEST_ID_TRACK_LIST_SECONDARY) {
        const tracks = msg.data as Array<{
          type: string;
          lang?: string;
          id: number;
        }>;
        if (Array.isArray(tracks)) {
          const { config } = loadConfig();
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
        if (mainWindow && !mainWindow.isDestroyed()) {
          tokenizeSubtitle(currentSubText).then((subtitleData) => {
            mainWindow!.webContents.send("subtitle:set", subtitleData);
          });
        }
      } else if (msg.request_id === MPV_REQUEST_ID_PATH) {
        updateCurrentMediaPath(msg.data);
      } else if (msg.request_id === MPV_REQUEST_ID_AID) {
        this.currentAudioTrackId =
          typeof msg.data === "number" ? (msg.data as number) : null;
        this.syncCurrentAudioStreamIndex();
      } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUBTEXT) {
        this.currentSecondarySubText = (msg.data as string) || "";
        if (mainWindow && !mainWindow.isDestroyed()) {
          mainWindow.webContents.send(
            "secondary-subtitle:set",
            this.currentSecondarySubText,
          );
        }
      } else if (msg.request_id === MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY) {
        previousSecondarySubVisibility =
          msg.data === true || msg.data === "yes";
        this.send({
          command: ["set_property", "secondary-sub-visibility", "no"],
        });
      }
    }
  }

  private autoLoadSecondarySubTrack(): void {
    const { config } = loadConfig();
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

  private subscribeToProperties(): void {
    this.send({ command: ["observe_property", 1, "sub-text"] });
    this.send({ command: ["observe_property", 2, "path"] });
    this.send({ command: ["observe_property", 3, "sub-start"] });
    this.send({ command: ["observe_property", 4, "sub-end"] });
    this.send({ command: ["observe_property", 5, "time-pos"] });
    this.send({ command: ["observe_property", 6, "secondary-sub-text"] });
    this.send({ command: ["observe_property", 7, "aid"] });
  }

  private getInitialState(): void {
    this.send({
      command: ["get_property", "sub-text"],
      request_id: MPV_REQUEST_ID_SUBTEXT,
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

  const normalizedText = text
    .replace(/\\N/g, " ")
    .replace(/\\n/g, " ")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!normalizedText) {
    return { text, tokens: null };
  }

  try {
    const rawTokens = await mecabTokenizer.tokenize(normalizedText);

    if (rawTokens && rawTokens.length > 0) {
      const mergedTokens = mergeTokens(rawTokens);
      return { text: normalizedText, tokens: mergedTokens };
    }
  } catch (err) {
    console.error("Tokenization error:", (err as Error).message);
  }

  return { text, tokens: null };
}

function updateOverlayBounds(geometry: WindowGeometry): void {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  if (!geometry) return;

  mainWindow.setBounds({
    x: geometry.x,
    y: geometry.y,
    width: geometry.width,
    height: geometry.height,
  });
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

function createMainWindow(): BrowserWindow {
  mainWindow = new BrowserWindow({
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
    },
  });

  const htmlPath = path.join(__dirname, "renderer", "index.html");
  console.log("Loading HTML from:", htmlPath);
  console.log("HTML file exists:", fs.existsSync(htmlPath));

  mainWindow.loadFile(htmlPath).catch((err) => {
    console.error("Failed to load HTML file:", err);
  });

  mainWindow.webContents.on(
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

  mainWindow.webContents.on("did-finish-load", () => {
    console.log("Overlay HTML loaded successfully");
  });

  mainWindow.webContents.on("before-input-event", (event, input) => {
    if (!overlayVisible) return;
    if (!shouldUseMarkAudioCardLocalFallback(input)) return;

    event.preventDefault();
    markLastCardAsAudioCard().catch((err) => {
      console.error("markLastCardAsAudioCard failed:", err);
      showMpvOsd(`Audio card failed: ${(err as Error).message}`);
    });
  });

  mainWindow.hide();

  mainWindow.on("closed", () => {
    mainWindow = null;
  });

  if (isDev) {
    mainWindow.webContents.openDevTools({ mode: "detach" });
  }

  return mainWindow;
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
  globalShortcut.register("Alt+Shift+O", () => {
    toggleOverlay();
  });

  globalShortcut.register("Alt+Shift+Y", () => {
    openYomitanSettings();
  });

  if (isDev) {
    globalShortcut.register("F12", () => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.toggleDevTools();
      }
    });
  }
}

function getConfiguredShortcuts() {
  const { config } = loadConfig();
  return {
    copySubtitle: config.shortcuts?.copySubtitle ?? "CommandOrControl+C",
    copySubtitleMultiple:
      config.shortcuts?.copySubtitleMultiple ?? "CommandOrControl+Shift+C",
    updateLastCardFromClipboard:
      config.shortcuts?.updateLastCardFromClipboard ?? "CommandOrControl+V",
    triggerFieldGrouping:
      config.shortcuts?.triggerFieldGrouping ?? "CommandOrControl+G",
    mineSentence: config.shortcuts?.mineSentence ?? "CommandOrControl+S",
    mineSentenceMultiple:
      config.shortcuts?.mineSentenceMultiple ?? "CommandOrControl+Shift+S",
    multiCopyTimeoutMs: config.shortcuts?.multiCopyTimeoutMs ?? 3000,
    toggleSecondarySub:
      config.shortcuts?.toggleSecondarySub ?? "CommandOrControl+Shift+V",
    markAudioCard:
      config.shortcuts?.markAudioCard ?? "CommandOrControl+Shift+A",
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

function showMpvOsd(text: string): void {
  if (mpvClient && mpvClient.connected && mpvClient.send) {
    mpvClient.send({
      command: ["show-text", text, "3000"],
    });
  } else {
    console.log("OSD (MPV not connected):", text);
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

ipcMain.on(
  "set-ignore-mouse-events",
  (
    _event: IpcMainEvent,
    ignore: boolean,
    options: { forward?: boolean } = {},
  ) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.setIgnoreMouseEvents(ignore, options);
    }
  },
);

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
  const { config } = loadConfig();
  if (config.ankiConnect?.autoUpdateNewCards !== false) {
    return;
  }

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
  if (shortcutsRegistered) return;

  const shortcuts = getConfiguredShortcuts();
  const { config } = loadConfig();
  const enableFieldGroupingShortcut =
    config.ankiConnect?.autoUpdateNewCards === false;

  if (shortcuts.copySubtitle) {
    globalShortcut.register(shortcuts.copySubtitle, () => {
      copyCurrentSubtitle();
    });
  }

  if (shortcuts.copySubtitleMultiple) {
    globalShortcut.register(shortcuts.copySubtitleMultiple, () => {
      startPendingMultiCopy(shortcuts.multiCopyTimeoutMs);
    });
  }

  if (shortcuts.updateLastCardFromClipboard) {
    globalShortcut.register(shortcuts.updateLastCardFromClipboard, () => {
      updateLastCardFromClipboard().catch((err) => {
        console.error("updateLastCardFromClipboard failed:", err);
        showMpvOsd(`Update failed: ${(err as Error).message}`);
      });
    });
  }

  if (enableFieldGroupingShortcut && shortcuts.triggerFieldGrouping) {
    globalShortcut.register(shortcuts.triggerFieldGrouping, () => {
      triggerFieldGrouping().catch((err) => {
        console.error("triggerFieldGrouping failed:", err);
        showMpvOsd(`Field grouping failed: ${(err as Error).message}`);
      });
    });
  }

  if (shortcuts.mineSentence) {
    globalShortcut.register(shortcuts.mineSentence, () => {
      mineSentenceCard().catch((err) => {
        console.error("mineSentenceCard failed:", err);
        showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
      });
    });
  }

  if (shortcuts.mineSentenceMultiple) {
    globalShortcut.register(shortcuts.mineSentenceMultiple, () => {
      startPendingMineSentenceMultiple(shortcuts.multiCopyTimeoutMs);
    });
  }

  if (shortcuts.toggleSecondarySub) {
    globalShortcut.register(shortcuts.toggleSecondarySub, () => {
      const cycle: SecondarySubMode[] = ["hidden", "visible", "hover"];
      const idx = cycle.indexOf(secondarySubMode);
      secondarySubMode = cycle[(idx + 1) % cycle.length];
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send(
          "secondary-subtitle:mode",
          secondarySubMode,
        );
      }
      showMpvOsd(`Secondary subtitle: ${secondarySubMode}`);
    });
  }

  if (shortcuts.markAudioCard) {
    globalShortcut.register(shortcuts.markAudioCard, () => {
      markLastCardAsAudioCard().catch((err) => {
        console.error("markLastCardAsAudioCard failed:", err);
        showMpvOsd(`Audio card failed: ${(err as Error).message}`);
      });
    });
  }

  shortcutsRegistered = true;
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

  shortcutsRegistered = false;
}

function updateOverlayVisibility(): void {
  console.log(
    "updateOverlayVisibility called, overlayVisible:",
    overlayVisible,
  );
  if (!mainWindow || mainWindow.isDestroyed()) {
    console.log("mainWindow not available");
    return;
  }

  if (!overlayVisible) {
    console.log("Hiding overlay");
    mainWindow.hide();
    unregisterOverlayShortcuts();

    if (
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
    }
  } else {
    console.log(
      "Should show overlay, isTracking:",
      windowTracker?.isTracking(),
    );

    if (mpvClient && mpvClient.connected) {
      mpvClient.send({
        command: ["get_property", "secondary-sub-visibility"],
        request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
      });
    }

    if (windowTracker && windowTracker.isTracking()) {
      const geometry = windowTracker.getGeometry();
      console.log("Geometry:", geometry);
      if (geometry) {
        updateOverlayBounds(geometry);
      }
      console.log("Showing mainWindow");
      mainWindow.show();
      mainWindow.focus();
      registerOverlayShortcuts();
    } else if (!windowTracker) {
      mainWindow.show();
      mainWindow.focus();
      registerOverlayShortcuts();
    }
  }
}

function setOverlayVisible(visible: boolean): void {
  overlayVisible = visible;
  updateOverlayVisibility();
  if (mpvClient && mpvClient.connected) {
    mpvClient.setSubVisibility(!visible);
  }
}

function toggleOverlay(): void {
  setOverlayVisible(!overlayVisible);
}

ipcMain.on(
  "set-ignore-mouse-events",
  (
    _event: IpcMainEvent,
    ignore: boolean,
    options: { forward?: boolean } = {},
  ) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.setIgnoreMouseEvents(ignore, options);
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
  return overlayVisible;
});

ipcMain.on("toggle-overlay", () => {
  toggleOverlay();
});

ipcMain.handle("get-current-subtitle", async () => {
  return await tokenizeSubtitle(currentSubText);
});

ipcMain.handle("get-subtitle-position", () => {
  return loadSubtitlePosition();
});

ipcMain.handle("get-subtitle-style", () => {
  const { config } = loadConfig();
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
    if (mpvClient && mpvClient.connected) {
      if (command[0] === "__replay-subtitle") {
        mpvClient.replayCurrentSubtitle();
      } else if (command[0] === "__play-next-subtitle") {
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

ipcMain.handle("get-anki-connect-status", () => {
  return ankiIntegration !== null;
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
      fieldGroupingResolver = resolve;
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send("kiku:field-grouping-request", data);
      } else {
        resolve({
          keepNoteId: 0,
          deleteNoteId: 0,
          deleteDuplicate: true,
          cancelled: true,
        });
        fieldGroupingResolver = null;
        return;
      }
      setTimeout(() => {
        if (fieldGroupingResolver) {
          fieldGroupingResolver({
            keepNoteId: 0,
            deleteNoteId: 0,
            deleteDuplicate: true,
            cancelled: true,
          });
          fieldGroupingResolver = null;
        }
      }, 90000);
    });
  };
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
    const { config } = loadConfig();
    if (!config.ankiConnect) {
      config.ankiConnect = {};
    }
    config.ankiConnect.enabled = enabled;
    saveConfig(config);

    if (enabled && !ankiIntegration && subtitleTimingTracker && mpvClient) {
      ankiIntegration = new AnkiIntegration(
        config.ankiConnect,
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
