import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { Command } from 'commander';
import { parse as parseJsonc } from 'jsonc-parser';
import { resolveConfigFilePath } from '../src/config/path-resolution.js';
import type {
  LogLevel,
  YoutubeSubgenMode,
  Backend,
  Args,
  LauncherYoutubeSubgenConfig,
  LauncherJellyfinConfig,
  PluginRuntimeConfig,
} from './types.js';
import {
  DEFAULT_SOCKET_PATH,
  DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS,
  DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS,
  DEFAULT_YOUTUBE_SUBGEN_OUT_DIR,
  DEFAULT_JIMAKU_API_BASE_URL,
} from './types.js';
import { log, fail } from './log.js';
import {
  resolvePathMaybe,
  isUrlTarget,
  uniqueNormalizedLangCodes,
  inferWhisperLanguage,
} from './util.js';

function resolveLauncherMainConfigPath(): string {
  return resolveConfigFilePath({
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
}

export function loadLauncherYoutubeSubgenConfig(): LauncherYoutubeSubgenConfig {
  const configPath = resolveLauncherMainConfigPath();
  if (!fs.existsSync(configPath)) return {};

  try {
    const data = fs.readFileSync(configPath, 'utf8');
    const parsed = configPath.endsWith('.jsonc') ? parseJsonc(data) : JSON.parse(data);
    if (!parsed || typeof parsed !== 'object') return {};
    const root = parsed as {
      youtubeSubgen?: unknown;
      secondarySub?: { secondarySubLanguages?: unknown };
      jimaku?: unknown;
    };
    const youtubeSubgen = root.youtubeSubgen;
    const mode =
      youtubeSubgen && typeof youtubeSubgen === 'object'
        ? (youtubeSubgen as { mode?: unknown }).mode
        : undefined;
    const whisperBin =
      youtubeSubgen && typeof youtubeSubgen === 'object'
        ? (youtubeSubgen as { whisperBin?: unknown }).whisperBin
        : undefined;
    const whisperModel =
      youtubeSubgen && typeof youtubeSubgen === 'object'
        ? (youtubeSubgen as { whisperModel?: unknown }).whisperModel
        : undefined;
    const primarySubLanguagesRaw =
      youtubeSubgen && typeof youtubeSubgen === 'object'
        ? (youtubeSubgen as { primarySubLanguages?: unknown }).primarySubLanguages
        : undefined;
    const secondarySubLanguagesRaw = root.secondarySub?.secondarySubLanguages;
    const primarySubLanguages = Array.isArray(primarySubLanguagesRaw)
      ? primarySubLanguagesRaw.filter((value): value is string => typeof value === 'string')
      : undefined;
    const secondarySubLanguages = Array.isArray(secondarySubLanguagesRaw)
      ? secondarySubLanguagesRaw.filter((value): value is string => typeof value === 'string')
      : undefined;
    const jimaku = root.jimaku;
    const jimakuApiKey =
      jimaku && typeof jimaku === 'object' ? (jimaku as { apiKey?: unknown }).apiKey : undefined;
    const jimakuApiKeyCommand =
      jimaku && typeof jimaku === 'object'
        ? (jimaku as { apiKeyCommand?: unknown }).apiKeyCommand
        : undefined;
    const jimakuApiBaseUrl =
      jimaku && typeof jimaku === 'object'
        ? (jimaku as { apiBaseUrl?: unknown }).apiBaseUrl
        : undefined;
    const jimakuLanguagePreference =
      jimaku && typeof jimaku === 'object'
        ? (jimaku as { languagePreference?: unknown }).languagePreference
        : undefined;
    const jimakuMaxEntryResults =
      jimaku && typeof jimaku === 'object'
        ? (jimaku as { maxEntryResults?: unknown }).maxEntryResults
        : undefined;
    const resolvedJimakuLanguagePreference =
      jimakuLanguagePreference === 'ja' ||
      jimakuLanguagePreference === 'en' ||
      jimakuLanguagePreference === 'none'
        ? jimakuLanguagePreference
        : undefined;
    const resolvedJimakuMaxEntryResults =
      typeof jimakuMaxEntryResults === 'number' &&
      Number.isFinite(jimakuMaxEntryResults) &&
      jimakuMaxEntryResults > 0
        ? Math.floor(jimakuMaxEntryResults)
        : undefined;

    return {
      mode: mode === 'automatic' || mode === 'preprocess' || mode === 'off' ? mode : undefined,
      whisperBin: typeof whisperBin === 'string' ? whisperBin : undefined,
      whisperModel: typeof whisperModel === 'string' ? whisperModel : undefined,
      primarySubLanguages,
      secondarySubLanguages,
      jimakuApiKey: typeof jimakuApiKey === 'string' ? jimakuApiKey : undefined,
      jimakuApiKeyCommand:
        typeof jimakuApiKeyCommand === 'string' ? jimakuApiKeyCommand : undefined,
      jimakuApiBaseUrl: typeof jimakuApiBaseUrl === 'string' ? jimakuApiBaseUrl : undefined,
      jimakuLanguagePreference: resolvedJimakuLanguagePreference,
      jimakuMaxEntryResults: resolvedJimakuMaxEntryResults,
    };
  } catch {
    return {};
  }
}

export function loadLauncherJellyfinConfig(): LauncherJellyfinConfig {
  const configPath = resolveLauncherMainConfigPath();
  if (!fs.existsSync(configPath)) return {};

  try {
    const data = fs.readFileSync(configPath, 'utf8');
    const parsed = configPath.endsWith('.jsonc') ? parseJsonc(data) : JSON.parse(data);
    if (!parsed || typeof parsed !== 'object') return {};
    const jellyfin = (parsed as { jellyfin?: unknown }).jellyfin;
    if (!jellyfin || typeof jellyfin !== 'object') return {};
    const typed = jellyfin as Record<string, unknown>;
    return {
      enabled: typeof typed.enabled === 'boolean' ? typed.enabled : undefined,
      serverUrl: typeof typed.serverUrl === 'string' ? typed.serverUrl : undefined,
      username: typeof typed.username === 'string' ? typed.username : undefined,
      accessToken: typeof typed.accessToken === 'string' ? typed.accessToken : undefined,
      userId: typeof typed.userId === 'string' ? typed.userId : undefined,
      defaultLibraryId:
        typeof typed.defaultLibraryId === 'string' ? typed.defaultLibraryId : undefined,
      pullPictures: typeof typed.pullPictures === 'boolean' ? typed.pullPictures : undefined,
      iconCacheDir: typeof typed.iconCacheDir === 'string' ? typed.iconCacheDir : undefined,
    };
  } catch {
    return {};
  }
}

function getPluginConfigCandidates(): string[] {
  const xdgConfigHome = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config');
  return Array.from(
    new Set([
      path.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
      path.join(os.homedir(), '.config', 'mpv', 'script-opts', 'subminer.conf'),
    ]),
  );
}

export function readPluginRuntimeConfig(logLevel: LogLevel): PluginRuntimeConfig {
  const runtimeConfig: PluginRuntimeConfig = {
    socketPath: DEFAULT_SOCKET_PATH,
  };
  const candidates = getPluginConfigCandidates();

  for (const configPath of candidates) {
    if (!fs.existsSync(configPath)) continue;
    try {
      const content = fs.readFileSync(configPath, 'utf8');
      const lines = content.split(/\r?\n/);
      for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.length === 0 || trimmed.startsWith('#')) continue;
        const socketMatch = trimmed.match(/^socket_path\s*=\s*(.+)$/i);
        if (socketMatch) {
          const value = (socketMatch[1] || '').split('#', 1)[0]?.trim() || '';
          if (value) runtimeConfig.socketPath = value;
        }
      }
      log(
        'debug',
        logLevel,
        `Using mpv plugin settings from ${configPath}: socket_path=${runtimeConfig.socketPath}`,
      );
      return runtimeConfig;
    } catch {
      log('warn', logLevel, `Failed to read ${configPath}; using launcher defaults`);
      return runtimeConfig;
    }
  }

  log(
    'debug',
    logLevel,
    `No mpv subminer.conf found; using launcher defaults (socket_path=${runtimeConfig.socketPath})`,
  );
  return runtimeConfig;
}

function ensureTarget(target: string, parsed: Args): void {
  if (isUrlTarget(target)) {
    parsed.target = target;
    parsed.targetKind = 'url';
    return;
  }
  const resolved = resolvePathMaybe(target);
  let stat: fs.Stats | null = null;
  try {
    stat = fs.statSync(resolved);
  } catch {
    stat = null;
  }
  if (stat?.isFile()) {
    parsed.target = resolved;
    parsed.targetKind = 'file';
    return;
  }
  if (stat?.isDirectory()) {
    parsed.directory = resolved;
    return;
  }
  fail(`Not a file, directory, or supported URL: ${target}`);
}

function parseLogLevel(value: string): LogLevel {
  if (value === 'debug' || value === 'info' || value === 'warn' || value === 'error') {
    return value;
  }
  fail(`Invalid log level: ${value} (must be debug, info, warn, or error)`);
}

function parseYoutubeMode(value: string): YoutubeSubgenMode {
  const normalized = value.toLowerCase();
  if (normalized === 'automatic' || normalized === 'preprocess' || normalized === 'off') {
    return normalized as YoutubeSubgenMode;
  }
  fail(`Invalid yt-subgen mode: ${value} (must be automatic, preprocess, or off)`);
}

function parseBackend(value: string): Backend {
  if (value === 'auto' || value === 'hyprland' || value === 'x11' || value === 'macos') {
    return value as Backend;
  }
  fail(`Invalid backend: ${value} (must be auto, hyprland, x11, or macos)`);
}

function applyRootOptions(program: Command): void {
  program
    .option('-b, --backend <backend>', 'Display backend')
    .option('-d, --directory <dir>', 'Directory to browse')
    .option('-r, --recursive', 'Search directories recursively')
    .option('-p, --profile <profile>', 'MPV profile')
    .option('--start', 'Explicitly start overlay')
    .option('--log-level <level>', 'Log level')
    .option('-R, --rofi', 'Use rofi picker')
    .option('-S, --start-overlay', 'Auto-start overlay')
    .option('-T, --no-texthooker', 'Disable texthooker-ui server');
}

function buildSubcommandHelpText(program: Command): string {
  const subcommands = program.commands
    .filter((command) => command.name() !== 'help')
    .map((command) => {
      const aliases = command.aliases();
      const term = aliases.length > 0 ? `${command.name()}|${aliases[0]}` : command.name();
      return { term, description: command.description() };
    });

  if (subcommands.length === 0) return '';
  const longestTerm = Math.max(...subcommands.map((entry) => entry.term.length));
  const lines = subcommands.map(
    (entry) => `  ${entry.term.padEnd(longestTerm)}  ${entry.description || ''}`.trimEnd(),
  );
  return `\nCommands:\n${lines.join('\n')}\n`;
}

function hasTopLevelCommand(argv: string[]): boolean {
  return getTopLevelCommand(argv) !== null;
}

function getTopLevelCommand(argv: string[]): { name: string; index: number } | null {
  const commandNames = new Set([
    'jellyfin',
    'jf',
    'yt',
    'youtube',
    'doctor',
    'config',
    'mpv',
    'texthooker',
    'app',
    'bin',
    'help',
  ]);
  const optionsWithValue = new Set([
    '-b',
    '--backend',
    '-d',
    '--directory',
    '-p',
    '--profile',
    '--log-level',
  ]);
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i] || '';
    if (token === '--') return null;
    if (token.startsWith('-')) {
      if (optionsWithValue.has(token)) {
        i += 1;
      }
      continue;
    }
    return commandNames.has(token) ? { name: token, index: i } : null;
  }
  return null;
}

export function parseArgs(
  argv: string[],
  scriptName: string,
  launcherConfig: LauncherYoutubeSubgenConfig,
): Args {
  const topLevelCommand = getTopLevelCommand(argv);
  const envMode = (process.env.SUBMINER_YT_SUBGEN_MODE || '').toLowerCase();
  const defaultMode: YoutubeSubgenMode =
    envMode === 'preprocess' || envMode === 'off' || envMode === 'automatic'
      ? (envMode as YoutubeSubgenMode)
      : launcherConfig.mode
        ? launcherConfig.mode
        : 'automatic';
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
  const youtubeAudioLangs = uniqueNormalizedLangCodes([...primarySubLangs, ...secondarySubLangs]);
  const parsed: Args = {
    backend: 'auto',
    directory: '.',
    recursive: false,
    profile: 'subminer',
    startOverlay: false,
    youtubeSubgenMode: defaultMode,
    whisperBin: process.env.SUBMINER_WHISPER_BIN || launcherConfig.whisperBin || '',
    whisperModel: process.env.SUBMINER_WHISPER_MODEL || launcherConfig.whisperModel || '',
    youtubeSubgenOutDir: process.env.SUBMINER_YT_SUBGEN_OUT_DIR || DEFAULT_YOUTUBE_SUBGEN_OUT_DIR,
    youtubeSubgenAudioFormat: process.env.SUBMINER_YT_SUBGEN_AUDIO_FORMAT || 'm4a',
    youtubeSubgenKeepTemp: process.env.SUBMINER_YT_SUBGEN_KEEP_TEMP === '1',
    jimakuApiKey: process.env.SUBMINER_JIMAKU_API_KEY || '',
    jimakuApiKeyCommand: process.env.SUBMINER_JIMAKU_API_KEY_COMMAND || '',
    jimakuApiBaseUrl: process.env.SUBMINER_JIMAKU_API_BASE_URL || DEFAULT_JIMAKU_API_BASE_URL,
    jimakuLanguagePreference: launcherConfig.jimakuLanguagePreference || 'ja',
    jimakuMaxEntryResults: launcherConfig.jimakuMaxEntryResults || 10,
    jellyfin: false,
    jellyfinLogin: false,
    jellyfinLogout: false,
    jellyfinPlay: false,
    jellyfinDiscovery: false,
    doctor: false,
    configPath: false,
    configShow: false,
    mpvIdle: false,
    mpvSocket: false,
    mpvStatus: false,
    appPassthrough: false,
    appArgs: [],
    jellyfinServer: '',
    jellyfinUsername: '',
    jellyfinPassword: '',
    youtubePrimarySubLangs: primarySubLangs,
    youtubeSecondarySubLangs: secondarySubLangs,
    youtubeAudioLangs,
    youtubeWhisperSourceLanguage: inferWhisperLanguage(primarySubLangs, 'ja'),
    useTexthooker: true,
    autoStartOverlay: false,
    texthookerOnly: false,
    useRofi: false,
    logLevel: 'info',
    target: '',
    targetKind: '',
  };

  if (launcherConfig.jimakuApiKey) parsed.jimakuApiKey = launcherConfig.jimakuApiKey;
  if (launcherConfig.jimakuApiKeyCommand)
    parsed.jimakuApiKeyCommand = launcherConfig.jimakuApiKeyCommand;
  if (launcherConfig.jimakuApiBaseUrl) parsed.jimakuApiBaseUrl = launcherConfig.jimakuApiBaseUrl;
  if (launcherConfig.jimakuLanguagePreference)
    parsed.jimakuLanguagePreference = launcherConfig.jimakuLanguagePreference;
  if (launcherConfig.jimakuMaxEntryResults !== undefined)
    parsed.jimakuMaxEntryResults = launcherConfig.jimakuMaxEntryResults;
  if (topLevelCommand && (topLevelCommand.name === 'app' || topLevelCommand.name === 'bin')) {
    parsed.appPassthrough = true;
    parsed.appArgs = argv.slice(topLevelCommand.index + 1);
    return parsed;
  }

  let jellyfinInvocation: {
    action?: string;
    discovery?: boolean;
    play?: boolean;
    login?: boolean;
    logout?: boolean;
    setup?: boolean;
    server?: string;
    username?: string;
    password?: string;
    logLevel?: string;
  } | null = null;
  let ytInvocation: {
    target?: string;
    mode?: string;
    outDir?: string;
    keepTemp?: boolean;
    whisperBin?: string;
    whisperModel?: string;
    ytSubgenAudioFormat?: string;
    logLevel?: string;
  } | null = null;
  let configInvocation: { action: string; logLevel?: string } | null = null;
  let mpvInvocation: { action: string; logLevel?: string } | null = null;
  let appInvocation: { appArgs: string[] } | null = null;
  let doctorLogLevel: string | null = null;
  let texthookerLogLevel: string | null = null;

  const commandProgram = new Command();
  commandProgram
    .name(scriptName)
    .description('Launch MPV with SubMiner sentence mining overlay')
    .showHelpAfterError(true)
    .enablePositionalOptions()
    .allowExcessArguments(false)
    .allowUnknownOption(false)
    .exitOverride();
  applyRootOptions(commandProgram);

  const rootProgram = new Command();
  rootProgram
    .name(scriptName)
    .description('Launch MPV with SubMiner sentence mining overlay')
    .usage('[options] [command] [target]')
    .showHelpAfterError(true)
    .allowExcessArguments(false)
    .allowUnknownOption(false)
    .exitOverride()
    .argument('[target]', 'file, directory, or URL');
  applyRootOptions(rootProgram);

  commandProgram
    .command('jellyfin')
    .alias('jf')
    .description('Jellyfin workflows')
    .argument('[action]', 'setup|discovery|play|login|logout')
    .option('-d, --discovery', 'Cast discovery mode')
    .option('-p, --play', 'Interactive play picker')
    .option('-l, --login', 'Login flow')
    .option('--logout', 'Clear token/session')
    .option('--setup', 'Open setup window')
    .option('-s, --server <url>', 'Jellyfin server URL')
    .option('-u, --username <name>', 'Jellyfin username')
    .option('-w, --password <pass>', 'Jellyfin password')
    .option('--log-level <level>', 'Log level')
    .action((action: string | undefined, options: Record<string, unknown>) => {
      jellyfinInvocation = {
        action,
        discovery: options.discovery === true,
        play: options.play === true,
        login: options.login === true,
        logout: options.logout === true,
        setup: options.setup === true,
        server: typeof options.server === 'string' ? options.server : undefined,
        username: typeof options.username === 'string' ? options.username : undefined,
        password: typeof options.password === 'string' ? options.password : undefined,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('yt')
    .alias('youtube')
    .description('YouTube workflows')
    .argument('[target]', 'YouTube URL or ytsearch: query')
    .option('-m, --mode <mode>', 'Subtitle generation mode')
    .option('-o, --out-dir <dir>', 'Subtitle output dir')
    .option('--keep-temp', 'Keep temp files')
    .option('--whisper-bin <path>', 'whisper.cpp CLI path')
    .option('--whisper-model <path>', 'whisper model path')
    .option('--yt-subgen-audio-format <format>', 'Audio extraction format')
    .option('--log-level <level>', 'Log level')
    .action((target: string | undefined, options: Record<string, unknown>) => {
      ytInvocation = {
        target,
        mode: typeof options.mode === 'string' ? options.mode : undefined,
        outDir: typeof options.outDir === 'string' ? options.outDir : undefined,
        keepTemp: options.keepTemp === true,
        whisperBin: typeof options.whisperBin === 'string' ? options.whisperBin : undefined,
        whisperModel: typeof options.whisperModel === 'string' ? options.whisperModel : undefined,
        ytSubgenAudioFormat:
          typeof options.ytSubgenAudioFormat === 'string' ? options.ytSubgenAudioFormat : undefined,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('doctor')
    .description('Run dependency and environment checks')
    .option('--log-level <level>', 'Log level')
    .action((options: Record<string, unknown>) => {
      parsed.doctor = true;
      doctorLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
    });

  commandProgram
    .command('config')
    .description('Config helpers')
    .argument('[action]', 'path|show', 'path')
    .option('--log-level <level>', 'Log level')
    .action((action: string, options: Record<string, unknown>) => {
      configInvocation = {
        action,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('mpv')
    .description('MPV helpers')
    .argument('[action]', 'status|socket|idle', 'status')
    .option('--log-level <level>', 'Log level')
    .action((action: string, options: Record<string, unknown>) => {
      mpvInvocation = {
        action,
        logLevel: typeof options.logLevel === 'string' ? options.logLevel : undefined,
      };
    });

  commandProgram
    .command('texthooker')
    .description('Launch texthooker-only mode')
    .option('--log-level <level>', 'Log level')
    .action((options: Record<string, unknown>) => {
      parsed.texthookerOnly = true;
      texthookerLogLevel = typeof options.logLevel === 'string' ? options.logLevel : null;
    });

  commandProgram
    .command('app')
    .alias('bin')
    .description('Pass arguments directly to SubMiner binary')
    .allowUnknownOption(true)
    .allowExcessArguments(true)
    .argument('[appArgs...]', 'Arguments forwarded to SubMiner app binary')
    .action((appArgs: string[] | undefined) => {
      appInvocation = {
        appArgs: Array.isArray(appArgs) ? appArgs : [],
      };
    });

  rootProgram.addHelpText('after', buildSubcommandHelpText(commandProgram));

  const selectedProgram = hasTopLevelCommand(argv) ? commandProgram : rootProgram;
  try {
    selectedProgram.parse(['node', scriptName, ...argv]);
  } catch (error) {
    const commanderError = error as { code?: string; message?: string };
    if (commanderError?.code === 'commander.helpDisplayed') {
      process.exit(0);
    }
    fail(commanderError?.message || String(error));
  }

  const options = selectedProgram.opts<Record<string, unknown>>();
  if (typeof options.backend === 'string') {
    parsed.backend = parseBackend(options.backend);
  }
  if (typeof options.directory === 'string') {
    parsed.directory = options.directory;
  }
  if (options.recursive === true) parsed.recursive = true;
  if (typeof options.profile === 'string') {
    parsed.profile = options.profile;
  }
  if (options.start === true) parsed.startOverlay = true;
  if (typeof options.logLevel === 'string') {
    parsed.logLevel = parseLogLevel(options.logLevel);
  }
  if (options.rofi === true) parsed.useRofi = true;
  if (options.startOverlay === true) parsed.autoStartOverlay = true;
  if (options.texthooker === false) parsed.useTexthooker = false;

  const rootTarget = rootProgram.processedArgs[0];
  if (typeof rootTarget === 'string' && rootTarget) {
    ensureTarget(rootTarget, parsed);
  }

  if (jellyfinInvocation) {
    if (jellyfinInvocation.logLevel) {
      parsed.logLevel = parseLogLevel(jellyfinInvocation.logLevel);
    }
    const action = (jellyfinInvocation.action || '').toLowerCase();
    if (action && !['setup', 'discovery', 'play', 'login', 'logout'].includes(action)) {
      fail(`Unknown jellyfin action: ${jellyfinInvocation.action}`);
    }

    parsed.jellyfinServer = jellyfinInvocation.server || '';
    parsed.jellyfinUsername = jellyfinInvocation.username || '';
    parsed.jellyfinPassword = jellyfinInvocation.password || '';

    const modeFlags = {
      setup: jellyfinInvocation.setup || action === 'setup',
      discovery: jellyfinInvocation.discovery || action === 'discovery',
      play: jellyfinInvocation.play || action === 'play',
      login: jellyfinInvocation.login || action === 'login',
      logout: jellyfinInvocation.logout || action === 'logout',
    };
    if (
      !modeFlags.setup &&
      !modeFlags.discovery &&
      !modeFlags.play &&
      !modeFlags.login &&
      !modeFlags.logout
    ) {
      modeFlags.setup = true;
    }

    parsed.jellyfin = Boolean(modeFlags.setup);
    parsed.jellyfinDiscovery = Boolean(modeFlags.discovery);
    parsed.jellyfinPlay = Boolean(modeFlags.play);
    parsed.jellyfinLogin = Boolean(modeFlags.login);
    parsed.jellyfinLogout = Boolean(modeFlags.logout);
  }

  if (ytInvocation) {
    if (ytInvocation.logLevel) {
      parsed.logLevel = parseLogLevel(ytInvocation.logLevel);
    }
    const mode = ytInvocation.mode;
    if (mode) parsed.youtubeSubgenMode = parseYoutubeMode(mode);
    const outDir = ytInvocation.outDir;
    if (outDir) parsed.youtubeSubgenOutDir = outDir;
    if (ytInvocation.keepTemp) {
      parsed.youtubeSubgenKeepTemp = true;
    }
    if (ytInvocation.whisperBin) parsed.whisperBin = ytInvocation.whisperBin;
    if (ytInvocation.whisperModel) parsed.whisperModel = ytInvocation.whisperModel;
    if (ytInvocation.ytSubgenAudioFormat) {
      parsed.youtubeSubgenAudioFormat = ytInvocation.ytSubgenAudioFormat;
    }
    if (ytInvocation.target) {
      ensureTarget(ytInvocation.target, parsed);
    }
  }

  if (doctorLogLevel) {
    parsed.logLevel = parseLogLevel(doctorLogLevel);
  }

  if (texthookerLogLevel) {
    parsed.logLevel = parseLogLevel(texthookerLogLevel);
  }

  if (configInvocation !== null) {
    if (configInvocation.logLevel) {
      parsed.logLevel = parseLogLevel(configInvocation.logLevel);
    }
    const action = (configInvocation.action || 'path').toLowerCase();
    if (action === 'path') parsed.configPath = true;
    else if (action === 'show') parsed.configShow = true;
    else fail(`Unknown config action: ${configInvocation.action}`);
  }

  if (mpvInvocation !== null) {
    if (mpvInvocation.logLevel) {
      parsed.logLevel = parseLogLevel(mpvInvocation.logLevel);
    }
    const action = (mpvInvocation.action || 'status').toLowerCase();
    if (action === 'status') parsed.mpvStatus = true;
    else if (action === 'socket') parsed.mpvSocket = true;
    else if (action === 'idle' || action === 'start') parsed.mpvIdle = true;
    else fail(`Unknown mpv action: ${mpvInvocation.action}`);
  }

  if (appInvocation !== null) {
    parsed.appPassthrough = true;
    parsed.appArgs = appInvocation.appArgs;
  }

  return parsed;
}
