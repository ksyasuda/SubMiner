/*
  mpv-yomitan - Yomitan integration for mpv
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
  shell,
  protocol,
  IpcMainEvent,
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
import * as os from "os";
import * as fs from "fs";
import WebSocket from "ws";
import { parse as parseJsonc } from "jsonc-parser";
import { MecabTokenizer } from "./mecab-tokenizer";
import { mergeTokens } from "./token-merger";
import { createWindowTracker, BaseWindowTracker } from "./window-trackers";
import {
  Config,
  SubtitleData,
  SubtitlePosition,
  Keybinding,
  WindowGeometry,
} from "./types";

const MPV_SOCKET_PATH = "/tmp/mpv-yomitan-socket";
const TEXTHOOKER_PORT = 5174;
const DEFAULT_WEBSOCKET_PORT = 6677;
let texthookerServer: http.Server | null = null;
let subtitleWebSocketServer: WebSocket.Server | null = null;
const USER_DATA_PATH = path.join(
  os.homedir(),
  ".config",
  "mpv-yomitan-overlay",
);
const isDev = process.argv.includes("--dev");

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
let mecabTokenizer: MecabTokenizer | null = null;
let keybindings: Keybinding[] = [];

const DEFAULT_KEYBINDINGS: Keybinding[] = [
  { key: "Space", command: ["cycle", "pause"] },
];

const CONFIG_DIR = path.join(
  os.homedir(),
  ".config",
  "mpv-yomitan-overlay",
);
const CONFIG_FILE_JSONC = path.join(CONFIG_DIR, "config.jsonc");
const CONFIG_FILE_JSON = path.join(CONFIG_DIR, "config.json");

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

function startTexthookerServer(): http.Server | null {
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

  texthookerServer.listen(TEXTHOOKER_PORT, "127.0.0.1", () => {
    console.log(
      `Texthooker server running at http://127.0.0.1:${TEXTHOOKER_PORT}`,
    );
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
  const mpvWebsocketPath = path.join(os.homedir(), ".config", "mpv", "mpv_websocket");
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

function loadSubtitlePosition(): SubtitlePosition | null {
  const { config } = loadConfig();
  const saved = config.subtitlePosition;

  if (saved && saved.yPercent !== undefined) {
    subtitlePosition = saved;
  } else {
    subtitlePosition = null;
  }

  return subtitlePosition;
}

function saveSubtitlePosition(position: SubtitlePosition): void {
  subtitlePosition = position;
  const { success, config } = loadConfig();
  if (!success) {
    console.error("Refusing to save - could not load existing config");
    return;
  }
  config.subtitlePosition = position;
  saveConfig(config);
}

function loadKeybindings(): Keybinding[] {
  const { config } = loadConfig();
  const userBindings = config.keybindings || [];

  const bindingMap = new Map<string, string[] | null>();

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

const isStartCommand = process.argv.includes("--start");
const isHelpCommand = process.argv.includes("--help");
const autoStartOverlay = process.argv.includes("--auto-start-overlay");

const gotTheLock = app.requestSingleInstanceLock();

if (!gotTheLock) {
  app.quit();
} else {
  app.on("second-instance", (_event, argv) => {
    handleCliCommand(argv);
  });

  if (!isStartCommand && !isHelpCommand) {
    console.error("No running instance. Use --start to launch the app.");
    app.quit();
  } else {
    app.whenReady().then(async () => {
      loadSubtitlePosition();
      loadKeybindings();

      mpvClient = new MpvIpcClient(MPV_SOCKET_PATH);
      mpvClient.connect();

      const { config } = loadConfig();
      const wsConfig = config.websocket || {};
      const wsEnabled = wsConfig.enabled ?? "auto";
      const wsPort = wsConfig.port || DEFAULT_WEBSOCKET_PORT;

      if (wsEnabled === true || (wsEnabled === "auto" && !hasMpvWebsocket())) {
        startSubtitleWebSocketServer(wsPort);
      } else if (wsEnabled === "auto") {
        console.log("mpv_websocket detected, skipping built-in WebSocket server");
      }

      mecabTokenizer = new MecabTokenizer();
      await mecabTokenizer.checkAvailability();

      await loadYomitanExtension();
      createMainWindow();
      registerGlobalShortcuts();

      windowTracker = createWindowTracker();
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
        };
        windowTracker.start();
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
    });

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createMainWindow();
      }
    });
  }
}

function handleCliCommand(argv: string[]): void {
  if (argv.includes("--stop")) {
    console.log("Stopping mpv-yomitan-overlay...");
    app.quit();
  } else if (argv.includes("--toggle")) {
    toggleOverlay();
  } else if (argv.includes("--settings") || argv.includes("--yomitan")) {
    setTimeout(() => {
      openYomitanSettings();
    }, 1000);
  } else if (argv.includes("--show")) {
    setOverlayVisible(true);
  } else if (argv.includes("--hide")) {
    setOverlayVisible(false);
  } else if (argv.includes("--texthooker")) {
    if (!texthookerServer) {
      startTexthookerServer();
    }
    const { config } = loadConfig();
    const openBrowser = config.texthooker?.openBrowser !== false;
    if (openBrowser) {
      shell.openExternal(`http://127.0.0.1:${TEXTHOOKER_PORT}`);
    }
    console.log(`Texthooker available at http://127.0.0.1:${TEXTHOOKER_PORT}`);
  } else if (argv.includes("--help")) {
    console.log(`
mpv-yomitan-overlay CLI commands:
  --start             Start the overlay app (required for first launch)
  --stop              Stop the running overlay app
  --toggle            Toggle subtitle overlay visibility
  --settings          Open Yomitan settings window
  --texthooker        Start texthooker server and open browser
  --show              Force show overlay
  --hide              Force hide overlay
  --auto-start-overlay  Auto-hide mpv subtitles on connect (show overlay)
  --dev               Run in development mode
  --help              Show this help
`);
    if (!mainWindow) app.quit();
  }
}

function handleInitialArgs(): void {
  handleCliCommand(process.argv);
}

interface MpvMessage {
  event?: string;
  name?: string;
  data?: unknown;
  request_id?: number;
}

class MpvIpcClient {
  private socketPath: string;
  public socket: net.Socket | null = null;
  private buffer = "";
  public connected = false;
  private reconnectAttempt = 0;
  private firstConnection = true;

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

      if (this.firstConnection && autoStartOverlay) {
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
      console.log(`Attempting to reconnect to MPV (attempt ${attempt + 1}, delay ${delay}ms)...`);
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
        broadcastSubtitle(currentSubText);
        if (mainWindow && !mainWindow.isDestroyed()) {
          const subtitleData = await tokenizeSubtitle(currentSubText);
          mainWindow.webContents.send("subtitle:set", subtitleData);
        }
      }
    } else if (msg.data !== undefined && msg.request_id) {
      if (msg.request_id === 101) {
        currentSubText = (msg.data as string) || "";
        broadcastSubtitle(currentSubText);
        if (mainWindow && !mainWindow.isDestroyed()) {
          tokenizeSubtitle(currentSubText).then((subtitleData) => {
            mainWindow!.webContents.send("subtitle:set", subtitleData);
          });
        }
      }
    }
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
  }

  private getInitialState(): void {
    this.send({ command: ["get_property", "sub-text"], request_id: 101 });
  }

  setSubVisibility(visible: boolean): void {
    this.send({ command: ["set_property", "sub-visibility", visible ? "yes" : "no"] });
  }
}


async function tokenizeSubtitle(text: string): Promise<SubtitleData> {
  if (!text || !mecabTokenizer) {
    return { text, tokens: null };
  }

  const normalizedText = text
    .replace(/\\N/g, ' ')
    .replace(/\\n/g, ' ')
    .replace(/\n/g, ' ')
    .replace(/\s+/g, ' ')
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

function updateOverlayVisibility(): void {
  console.log("updateOverlayVisibility called, overlayVisible:", overlayVisible);
  if (!mainWindow || mainWindow.isDestroyed()) {
    console.log("mainWindow not available");
    return;
  }

  if (!overlayVisible) {
    console.log("Hiding overlay");
    mainWindow.hide();
  } else {
    console.log(
      "Should show overlay, isTracking:",
      windowTracker?.isTracking(),
    );
    if (windowTracker && windowTracker.isTracking()) {
      const geometry = windowTracker.getGeometry();
      console.log("Geometry:", geometry);
      if (geometry) {
        updateOverlayBounds(geometry);
      }
      console.log("Showing mainWindow");
      mainWindow.show();
      mainWindow.focus();
    } else if (!windowTracker) {
      mainWindow.show();
      mainWindow.focus();
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

function ensureExtensionCopy(sourceDir: string): string {
  if (process.platform !== "linux") {
    return sourceDir;
  }

  const extensionsRoot = path.join(USER_DATA_PATH, "extensions");
  const targetDir = path.join(extensionsRoot, "yomitan");

  const sourceManifest = path.join(sourceDir, "manifest.json");
  const targetManifest = path.join(targetDir, "manifest.json");

  let shouldCopy = !fs.existsSync(targetDir);
  if (!shouldCopy && fs.existsSync(sourceManifest) && fs.existsSync(targetManifest)) {
    try {
      const sourceVersion = (JSON.parse(fs.readFileSync(sourceManifest, "utf-8")) as { version: string }).version;
      const targetVersion = (JSON.parse(fs.readFileSync(targetManifest, "utf-8")) as { version: string }).version;
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
    "/usr/share/mpv-yomitan-overlay/yomitan",
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

  mainWindow.webContents.on("did-fail-load", (_event, errorCode, errorDescription, validatedURL) => {
    console.error("Page failed to load:", errorCode, errorDescription, validatedURL);
  });

  mainWindow.webContents.on("did-finish-load", () => {
    console.log("Overlay HTML loaded successfully");
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
    yomitanSettingsWindow!.loadURL(settingsUrl).then(() => {
      console.log("Settings URL loaded successfully");
    }).catch((err: Error) => {
      console.error("Failed to load settings URL:", err);
      loadAttempts++;
      if (loadAttempts < maxAttempts && yomitanSettingsWindow && !yomitanSettingsWindow.isDestroyed()) {
        console.log(`Retrying in 500ms (attempt ${loadAttempts + 1}/${maxAttempts})`);
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

ipcMain.on("set-ignore-mouse-events", (_event: IpcMainEvent, ignore: boolean, options: { forward?: boolean } = {}) => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.setIgnoreMouseEvents(ignore, options);
  }
});

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
  return subtitlePosition;
});

ipcMain.on("save-subtitle-position", (_event: IpcMainEvent, position: SubtitlePosition) => {
  saveSubtitlePosition(position);
});

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

ipcMain.on("mpv-command", (_event: IpcMainEvent, command: string[]) => {
  if (mpvClient && mpvClient.connected) {
    mpvClient.send({ command });
  }
});

ipcMain.handle("get-keybindings", () => {
  return keybindings;
});

