import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { parse as parseJsonc } from "jsonc-parser";
import type {
  LogLevel, YoutubeSubgenMode, Backend, Args,
  LauncherYoutubeSubgenConfig, LauncherJellyfinConfig, PluginRuntimeConfig,
} from "./types.js";
import {
  DEFAULT_SOCKET_PATH, DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS,
  DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS, DEFAULT_YOUTUBE_SUBGEN_OUT_DIR,
  DEFAULT_JIMAKU_API_BASE_URL,
} from "./types.js";
import { log, fail } from "./log.js";
import {
  resolvePathMaybe, isUrlTarget, uniqueNormalizedLangCodes, parseBoolLike,
  inferWhisperLanguage,
} from "./util.js";

export function usage(scriptName: string): string {
  return `subminer - Launch MPV with SubMiner sentence mining overlay

Usage: ${scriptName} [OPTIONS] [FILE|DIRECTORY|URL]

Options:
  -b, --backend BACKEND  Display backend to use: auto, hyprland, x11, macos (default: auto)
  -d, --directory DIR    Directory to browse for videos (default: current directory)
  -r, --recursive        Search for videos recursively
  -p, --profile PROFILE  MPV profile to use (default: subminer)
      --start            Explicitly start SubMiner overlay
      --yt-subgen-mode MODE
                         YouTube subtitle generation mode: automatic, preprocess, off (default: automatic)
      --whisper-bin PATH whisper.cpp CLI binary (used for fallback transcription)
      --whisper-model PATH
                         whisper model file path (used for fallback transcription)
      --yt-subgen-out-dir DIR
                         Output directory for generated YouTube subtitles (default: ${DEFAULT_YOUTUBE_SUBGEN_OUT_DIR})
      --yt-subgen-audio-format FORMAT
                         Audio format for extraction (default: m4a)
      --yt-subgen-keep-temp
                         Keep YouTube subtitle temp directory
       --log-level LEVEL  Set log level: debug, info, warn, error
  -R, --rofi             Use rofi file browser instead of fzf for video selection
  -S, --start-overlay    Auto-start SubMiner overlay after MPV socket is ready
  -T, --no-texthooker    Disable texthooker-ui server
      --texthooker       Launch only texthooker page (no MPV/overlay workflow)
      --jellyfin         Open Jellyfin setup window in the app
      --jellyfin-login   Login via app CLI (requires server/username/password)
      --jellyfin-logout  Clear Jellyfin token/session via app CLI
      --jellyfin-play    Pick Jellyfin library/item and play in mpv
      --jellyfin-server URL
                         Jellyfin server URL (for login/play menu API)
      --jellyfin-username NAME
                         Jellyfin username (for login)
      --jellyfin-password PASS
                         Jellyfin password (for login)
  -h, --help             Show this help message

Environment:
  SUBMINER_APPIMAGE_PATH      Path to SubMiner AppImage/binary (optional override)
  SUBMINER_ROFI_THEME         Path to rofi theme file (optional override)
  SUBMINER_YT_SUBGEN_MODE     automatic, preprocess, off (optional default)
  SUBMINER_WHISPER_BIN        whisper.cpp binary path (optional fallback)
  SUBMINER_WHISPER_MODEL      whisper model path (optional fallback)
  SUBMINER_YT_SUBGEN_OUT_DIR  Generated subtitle output directory

Examples:
  ${scriptName}                        # Browse current directory with fzf
  ${scriptName} -R                     # Browse current directory with rofi
  ${scriptName} -d ~/Videos            # Browse ~/Videos
  ${scriptName} -r -d ~/Anime          # Recursively browse ~/Anime
  ${scriptName} video.mkv              # Play specific file
  ${scriptName} https://youtu.be/...   # Play a YouTube URL
  ${scriptName} ytsearch:query         # Play first YouTube search result
  ${scriptName} --yt-subgen-mode preprocess --whisper-bin /path/whisper-cli --whisper-model /path/model.bin https://youtu.be/...
  ${scriptName} video.mkv              # Play with subminer profile
  ${scriptName} -p gpu-hq video.mkv    # Play with gpu-hq profile
  ${scriptName} -b x11 video.mkv       # Force x11 backend
  ${scriptName} -S video.mkv           # Start overlay immediately after MPV launch
  ${scriptName} --texthooker           # Launch only texthooker page
  ${scriptName} --jellyfin             # Open Jellyfin setup form window
  ${scriptName} --jellyfin-play        # Pick Jellyfin library/item and play
`;
}

export function loadLauncherYoutubeSubgenConfig(): LauncherYoutubeSubgenConfig {
  const configDir = path.join(os.homedir(), ".config", "SubMiner");
  const jsoncPath = path.join(configDir, "config.jsonc");
  const jsonPath = path.join(configDir, "config.json");
  const configPath = fs.existsSync(jsoncPath)
    ? jsoncPath
    : fs.existsSync(jsonPath)
      ? jsonPath
      : "";
  if (!configPath) return {};

  try {
    const data = fs.readFileSync(configPath, "utf8");
    const parsed = configPath.endsWith(".jsonc")
      ? parseJsonc(data)
      : JSON.parse(data);
    if (!parsed || typeof parsed !== "object") return {};
    const root = parsed as {
      youtubeSubgen?: unknown;
      secondarySub?: { secondarySubLanguages?: unknown };
      jimaku?: unknown;
    };
    const youtubeSubgen = root.youtubeSubgen;
    const mode =
      youtubeSubgen && typeof youtubeSubgen === "object"
        ? (youtubeSubgen as { mode?: unknown }).mode
        : undefined;
    const whisperBin =
      youtubeSubgen && typeof youtubeSubgen === "object"
        ? (youtubeSubgen as { whisperBin?: unknown }).whisperBin
        : undefined;
    const whisperModel =
      youtubeSubgen && typeof youtubeSubgen === "object"
        ? (youtubeSubgen as { whisperModel?: unknown }).whisperModel
        : undefined;
    const primarySubLanguagesRaw =
      youtubeSubgen && typeof youtubeSubgen === "object"
        ? (youtubeSubgen as { primarySubLanguages?: unknown }).primarySubLanguages
        : undefined;
    const secondarySubLanguagesRaw = root.secondarySub?.secondarySubLanguages;
    const primarySubLanguages = Array.isArray(primarySubLanguagesRaw)
      ? primarySubLanguagesRaw.filter(
          (value): value is string => typeof value === "string",
        )
      : undefined;
    const secondarySubLanguages = Array.isArray(secondarySubLanguagesRaw)
      ? secondarySubLanguagesRaw.filter(
          (value): value is string => typeof value === "string",
        )
      : undefined;
    const jimaku = root.jimaku;
    const jimakuApiKey =
      jimaku && typeof jimaku === "object"
        ? (jimaku as { apiKey?: unknown }).apiKey
        : undefined;
    const jimakuApiKeyCommand =
      jimaku && typeof jimaku === "object"
        ? (jimaku as { apiKeyCommand?: unknown }).apiKeyCommand
        : undefined;
    const jimakuApiBaseUrl =
      jimaku && typeof jimaku === "object"
        ? (jimaku as { apiBaseUrl?: unknown }).apiBaseUrl
        : undefined;
    const jimakuLanguagePreference = jimaku && typeof jimaku === "object"
      ? (jimaku as { languagePreference?: unknown }).languagePreference
      : undefined;
    const jimakuMaxEntryResults =
      jimaku && typeof jimaku === "object"
        ? (jimaku as { maxEntryResults?: unknown }).maxEntryResults
        : undefined;
    const resolvedJimakuLanguagePreference =
      jimakuLanguagePreference === "ja" ||
      jimakuLanguagePreference === "en" ||
      jimakuLanguagePreference === "none"
        ? jimakuLanguagePreference
        : undefined;
    const resolvedJimakuMaxEntryResults =
      typeof jimakuMaxEntryResults === "number" &&
      Number.isFinite(jimakuMaxEntryResults) &&
      jimakuMaxEntryResults > 0
        ? Math.floor(jimakuMaxEntryResults)
        : undefined;

    return {
      mode:
        mode === "automatic" || mode === "preprocess" || mode === "off"
          ? mode
          : undefined,
      whisperBin: typeof whisperBin === "string" ? whisperBin : undefined,
      whisperModel: typeof whisperModel === "string" ? whisperModel : undefined,
      primarySubLanguages,
      secondarySubLanguages,
      jimakuApiKey: typeof jimakuApiKey === "string" ? jimakuApiKey : undefined,
      jimakuApiKeyCommand:
        typeof jimakuApiKeyCommand === "string"
          ? jimakuApiKeyCommand
          : undefined,
      jimakuApiBaseUrl:
        typeof jimakuApiBaseUrl === "string"
          ? jimakuApiBaseUrl
          : undefined,
      jimakuLanguagePreference: resolvedJimakuLanguagePreference,
      jimakuMaxEntryResults: resolvedJimakuMaxEntryResults,
    };
  } catch {
    return {};
  }
}

export function loadLauncherJellyfinConfig(): LauncherJellyfinConfig {
  const configDir = path.join(os.homedir(), ".config", "SubMiner");
  const jsoncPath = path.join(configDir, "config.jsonc");
  const jsonPath = path.join(configDir, "config.json");
  const configPath = fs.existsSync(jsoncPath)
    ? jsoncPath
    : fs.existsSync(jsonPath)
      ? jsonPath
      : "";
  if (!configPath) return {};

  try {
    const data = fs.readFileSync(configPath, "utf8");
    const parsed = configPath.endsWith(".jsonc")
      ? parseJsonc(data)
      : JSON.parse(data);
    if (!parsed || typeof parsed !== "object") return {};
    const jellyfin = (parsed as { jellyfin?: unknown }).jellyfin;
    if (!jellyfin || typeof jellyfin !== "object") return {};
    const typed = jellyfin as Record<string, unknown>;
    return {
      enabled:
        typeof typed.enabled === "boolean" ? typed.enabled : undefined,
      serverUrl:
        typeof typed.serverUrl === "string" ? typed.serverUrl : undefined,
      username:
        typeof typed.username === "string" ? typed.username : undefined,
      accessToken:
        typeof typed.accessToken === "string" ? typed.accessToken : undefined,
      userId:
        typeof typed.userId === "string" ? typed.userId : undefined,
      defaultLibraryId:
        typeof typed.defaultLibraryId === "string"
          ? typed.defaultLibraryId
          : undefined,
      pullPictures:
        typeof typed.pullPictures === "boolean"
          ? typed.pullPictures
          : undefined,
      iconCacheDir:
        typeof typed.iconCacheDir === "string"
          ? typed.iconCacheDir
          : undefined,
    };
  } catch {
    return {};
  }
}

function getPluginConfigCandidates(): string[] {
  const xdgConfigHome =
    process.env.XDG_CONFIG_HOME || path.join(os.homedir(), ".config");
  return Array.from(
    new Set([
      path.join(xdgConfigHome, "mpv", "script-opts", "subminer.conf"),
      path.join(os.homedir(), ".config", "mpv", "script-opts", "subminer.conf"),
    ]),
  );
}

export function readPluginRuntimeConfig(logLevel: LogLevel): PluginRuntimeConfig {
  const runtimeConfig: PluginRuntimeConfig = {
    autoStartOverlay: false,
    socketPath: DEFAULT_SOCKET_PATH,
  };
  const candidates = getPluginConfigCandidates();

  for (const configPath of candidates) {
    if (!fs.existsSync(configPath)) continue;
    try {
      const content = fs.readFileSync(configPath, "utf8");
      const lines = content.split(/\r?\n/);
      for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.length === 0 || trimmed.startsWith("#")) continue;
        const autoStartMatch = trimmed.match(/^auto_start\s*=\s*(.+)$/i);
        if (autoStartMatch) {
          const value = (autoStartMatch[1] || "").split("#", 1)[0]?.trim() || "";
          const parsed = parseBoolLike(value);
          if (parsed !== null) {
            runtimeConfig.autoStartOverlay = parsed;
          }
          continue;
        }

        const socketMatch = trimmed.match(/^socket_path\s*=\s*(.+)$/i);
        if (socketMatch) {
          const value = (socketMatch[1] || "").split("#", 1)[0]?.trim() || "";
          if (value) runtimeConfig.socketPath = value;
        }
      }
      log(
        "debug",
        logLevel,
        `Using mpv plugin settings from ${configPath}: auto_start=${runtimeConfig.autoStartOverlay ? "yes" : "no"} socket_path=${runtimeConfig.socketPath}`,
      );
      return runtimeConfig;
    } catch {
      log(
        "warn",
        logLevel,
        `Failed to read ${configPath}; using launcher defaults`,
      );
      return runtimeConfig;
    }
  }

  log(
    "debug",
    logLevel,
    `No mpv subminer.conf found; using launcher defaults (auto_start=no socket_path=${runtimeConfig.socketPath})`,
  );
  return runtimeConfig;
}

export function parseArgs(
  argv: string[],
  scriptName: string,
  launcherConfig: LauncherYoutubeSubgenConfig,
): Args {
  const envMode = (process.env.SUBMINER_YT_SUBGEN_MODE || "").toLowerCase();
  const defaultMode: YoutubeSubgenMode =
    envMode === "preprocess" || envMode === "off" || envMode === "automatic"
      ? (envMode as YoutubeSubgenMode)
      : launcherConfig.mode
        ? launcherConfig.mode
        : "automatic";
  const configuredSecondaryLangs = uniqueNormalizedLangCodes(
    launcherConfig.secondarySubLanguages ?? [],
  );
  const configuredPrimaryLangs = uniqueNormalizedLangCodes(
    launcherConfig.primarySubLanguages ?? [],
  );
  const primarySubLangs =
    configuredPrimaryLangs.length > 0
      ? configuredPrimaryLangs
      : [...DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS];
  const secondarySubLangs =
    configuredSecondaryLangs.length > 0
      ? configuredSecondaryLangs
      : [...DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS];
  const youtubeAudioLangs = uniqueNormalizedLangCodes([
    ...primarySubLangs,
    ...secondarySubLangs,
  ]);
  const parsed: Args = {
    backend: "auto",
    directory: ".",
    recursive: false,
    profile: "subminer",
    startOverlay: false,
    youtubeSubgenMode: defaultMode,
    whisperBin: process.env.SUBMINER_WHISPER_BIN || launcherConfig.whisperBin || "",
    whisperModel:
      process.env.SUBMINER_WHISPER_MODEL || launcherConfig.whisperModel || "",
    youtubeSubgenOutDir:
      process.env.SUBMINER_YT_SUBGEN_OUT_DIR || DEFAULT_YOUTUBE_SUBGEN_OUT_DIR,
    youtubeSubgenAudioFormat: process.env.SUBMINER_YT_SUBGEN_AUDIO_FORMAT || "m4a",
    youtubeSubgenKeepTemp: process.env.SUBMINER_YT_SUBGEN_KEEP_TEMP === "1",
    jimakuApiKey: process.env.SUBMINER_JIMAKU_API_KEY || "",
    jimakuApiKeyCommand: process.env.SUBMINER_JIMAKU_API_KEY_COMMAND || "",
    jimakuApiBaseUrl:
      process.env.SUBMINER_JIMAKU_API_BASE_URL || DEFAULT_JIMAKU_API_BASE_URL,
    jimakuLanguagePreference:
      launcherConfig.jimakuLanguagePreference || "ja",
    jimakuMaxEntryResults: launcherConfig.jimakuMaxEntryResults || 10,
    jellyfin: false,
    jellyfinLogin: false,
    jellyfinLogout: false,
    jellyfinPlay: false,
    jellyfinServer: "",
    jellyfinUsername: "",
    jellyfinPassword: "",
    youtubePrimarySubLangs: primarySubLangs,
    youtubeSecondarySubLangs: secondarySubLangs,
    youtubeAudioLangs,
    youtubeWhisperSourceLanguage: inferWhisperLanguage(primarySubLangs, "ja"),
    useTexthooker: true,
    autoStartOverlay: false,
    texthookerOnly: false,
    useRofi: false,
    logLevel: "info",
    target: "",
    targetKind: "",
  };

  if (launcherConfig.jimakuApiKey) parsed.jimakuApiKey = launcherConfig.jimakuApiKey;
  if (launcherConfig.jimakuApiKeyCommand)
    parsed.jimakuApiKeyCommand = launcherConfig.jimakuApiKeyCommand;
  if (launcherConfig.jimakuApiBaseUrl)
    parsed.jimakuApiBaseUrl = launcherConfig.jimakuApiBaseUrl;
  if (launcherConfig.jimakuLanguagePreference)
    parsed.jimakuLanguagePreference = launcherConfig.jimakuLanguagePreference;
  if (launcherConfig.jimakuMaxEntryResults !== undefined)
    parsed.jimakuMaxEntryResults = launcherConfig.jimakuMaxEntryResults;

  const isValidLogLevel = (value: string): value is LogLevel =>
    value === "debug" ||
    value === "info" ||
    value === "warn" ||
    value === "error";
  const isValidYoutubeSubgenMode = (value: string): value is YoutubeSubgenMode =>
    value === "automatic" || value === "preprocess" || value === "off";

  let i = 0;
  while (i < argv.length) {
    const arg = argv[i];

    if (arg === "-b" || arg === "--backend") {
      const value = argv[i + 1];
      if (!value) fail("--backend requires a value");
      if (!["auto", "hyprland", "x11", "macos"].includes(value)) {
        fail(
          `Invalid backend: ${value} (must be auto, hyprland, x11, or macos)`,
        );
      }
      parsed.backend = value as Backend;
      i += 2;
      continue;
    }

    if (arg === "-d" || arg === "--directory") {
      const value = argv[i + 1];
      if (!value) fail("--directory requires a value");
      parsed.directory = value;
      i += 2;
      continue;
    }

    if (arg === "-r" || arg === "--recursive") {
      parsed.recursive = true;
      i += 1;
      continue;
    }

    if (arg === "-p" || arg === "--profile") {
      const value = argv[i + 1];
      if (!value) fail("--profile requires a value");
      parsed.profile = value;
      i += 2;
      continue;
    }

    if (arg === "--start") {
      parsed.startOverlay = true;
      i += 1;
      continue;
    }

    if (arg === "--yt-subgen-mode" || arg === "--youtube-subgen-mode") {
      const value = (argv[i + 1] || "").toLowerCase();
      if (!isValidYoutubeSubgenMode(value)) {
        fail(
          `Invalid yt-subgen mode: ${value || ""} (must be automatic, preprocess, or off)`,
        );
      }
      parsed.youtubeSubgenMode = value;
      i += 2;
      continue;
    }

    if (
      arg.startsWith("--yt-subgen-mode=") ||
      arg.startsWith("--youtube-subgen-mode=")
    ) {
      const value = arg.split("=", 2)[1]?.toLowerCase() || "";
      if (!isValidYoutubeSubgenMode(value)) {
        fail(
          `Invalid yt-subgen mode: ${value || ""} (must be automatic, preprocess, or off)`,
        );
      }
      parsed.youtubeSubgenMode = value;
      i += 1;
      continue;
    }

    if (arg === "--whisper-bin") {
      const value = argv[i + 1];
      if (!value) fail("--whisper-bin requires a value");
      parsed.whisperBin = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--whisper-bin=")) {
      const value = arg.slice("--whisper-bin=".length);
      if (!value) fail("--whisper-bin requires a value");
      parsed.whisperBin = value;
      i += 1;
      continue;
    }

    if (arg === "--whisper-model") {
      const value = argv[i + 1];
      if (!value) fail("--whisper-model requires a value");
      parsed.whisperModel = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--whisper-model=")) {
      const value = arg.slice("--whisper-model=".length);
      if (!value) fail("--whisper-model requires a value");
      parsed.whisperModel = value;
      i += 1;
      continue;
    }

    if (arg === "--yt-subgen-out-dir") {
      const value = argv[i + 1];
      if (!value) fail("--yt-subgen-out-dir requires a value");
      parsed.youtubeSubgenOutDir = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--yt-subgen-out-dir=")) {
      const value = arg.slice("--yt-subgen-out-dir=".length);
      if (!value) fail("--yt-subgen-out-dir requires a value");
      parsed.youtubeSubgenOutDir = value;
      i += 1;
      continue;
    }

    if (arg === "--yt-subgen-audio-format") {
      const value = argv[i + 1];
      if (!value) fail("--yt-subgen-audio-format requires a value");
      parsed.youtubeSubgenAudioFormat = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--yt-subgen-audio-format=")) {
      const value = arg.slice("--yt-subgen-audio-format=".length);
      if (!value) fail("--yt-subgen-audio-format requires a value");
      parsed.youtubeSubgenAudioFormat = value;
      i += 1;
      continue;
    }

    if (arg === "--yt-subgen-keep-temp") {
      parsed.youtubeSubgenKeepTemp = true;
      i += 1;
      continue;
    }

    if (arg === "--log-level") {
      const value = argv[i + 1];
      if (!value || !isValidLogLevel(value)) {
        fail(
          `Invalid log level: ${value ?? ""} (must be debug, info, warn, or error)`,
        );
      }
      parsed.logLevel = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--log-level=")) {
      const value = arg.slice("--log-level=".length);
      if (!isValidLogLevel(value)) {
        fail(
          `Invalid log level: ${value} (must be debug, info, warn, or error)`,
        );
      }
      parsed.logLevel = value;
      i += 1;
      continue;
    }

    if (arg === "-R" || arg === "--rofi") {
      parsed.useRofi = true;
      i += 1;
      continue;
    }

    if (arg === "-S" || arg === "--start-overlay") {
      parsed.autoStartOverlay = true;
      i += 1;
      continue;
    }

    if (arg === "-T" || arg === "--no-texthooker") {
      parsed.useTexthooker = false;
      i += 1;
      continue;
    }

    if (arg === "--texthooker") {
      parsed.texthookerOnly = true;
      i += 1;
      continue;
    }

    if (arg === "--jellyfin") {
      parsed.jellyfin = true;
      i += 1;
      continue;
    }

    if (arg === "--jellyfin-login") {
      parsed.jellyfinLogin = true;
      i += 1;
      continue;
    }

    if (arg === "--jellyfin-logout") {
      parsed.jellyfinLogout = true;
      i += 1;
      continue;
    }

    if (arg === "--jellyfin-play") {
      parsed.jellyfinPlay = true;
      i += 1;
      continue;
    }

    if (arg === "--jellyfin-server") {
      const value = argv[i + 1];
      if (!value) fail("--jellyfin-server requires a value");
      parsed.jellyfinServer = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--jellyfin-server=")) {
      parsed.jellyfinServer = arg.split("=", 2)[1] || "";
      i += 1;
      continue;
    }

    if (arg === "--jellyfin-username") {
      const value = argv[i + 1];
      if (!value) fail("--jellyfin-username requires a value");
      parsed.jellyfinUsername = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--jellyfin-username=")) {
      parsed.jellyfinUsername = arg.split("=", 2)[1] || "";
      i += 1;
      continue;
    }

    if (arg === "--jellyfin-password") {
      const value = argv[i + 1];
      if (!value) fail("--jellyfin-password requires a value");
      parsed.jellyfinPassword = value;
      i += 2;
      continue;
    }

    if (arg.startsWith("--jellyfin-password=")) {
      parsed.jellyfinPassword = arg.split("=", 2)[1] || "";
      i += 1;
      continue;
    }

    if (arg === "-h" || arg === "--help") {
      process.stdout.write(usage(scriptName));
      process.exit(0);
    }

    if (arg === "--") {
      i += 1;
      break;
    }

    if (arg.startsWith("-")) {
      fail(`Unknown option: ${arg}`);
    }

    break;
  }

  const positional = argv.slice(i);
  if (positional.length > 0) {
    const target = positional[0];
    if (isUrlTarget(target)) {
      parsed.target = target;
      parsed.targetKind = "url";
    } else {
      const resolved = resolvePathMaybe(target);
      if (fs.existsSync(resolved) && fs.statSync(resolved).isFile()) {
        parsed.target = resolved;
        parsed.targetKind = "file";
      } else if (
        fs.existsSync(resolved) &&
        fs.statSync(resolved).isDirectory()
      ) {
        parsed.directory = resolved;
      } else {
        fail(`Not a file, directory, or supported URL: ${target}`);
      }
    }
  }

  return parsed;
}
