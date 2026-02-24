import fs from 'node:fs';
import { fail } from '../log.js';
import type {
  Args,
  Backend,
  LauncherYoutubeSubgenConfig,
  LogLevel,
  YoutubeSubgenMode,
} from '../types.js';
import {
  DEFAULT_JIMAKU_API_BASE_URL,
  DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS,
  DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS,
  DEFAULT_YOUTUBE_SUBGEN_OUT_DIR,
} from '../types.js';
import {
  inferWhisperLanguage,
  isUrlTarget,
  resolvePathMaybe,
  uniqueNormalizedLangCodes,
} from '../util.js';
import type { CliInvocations } from './cli-parser-builder.js';

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

export function createDefaultArgs(launcherConfig: LauncherYoutubeSubgenConfig): Args {
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
    passwordStore: '',
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

  return parsed;
}

export function applyRootOptionsToArgs(
  parsed: Args,
  options: Record<string, unknown>,
  rootTarget: unknown,
): void {
  if (typeof options.backend === 'string') parsed.backend = parseBackend(options.backend);
  if (typeof options.directory === 'string') parsed.directory = options.directory;
  if (options.recursive === true) parsed.recursive = true;
  if (typeof options.profile === 'string') parsed.profile = options.profile;
  if (options.start === true) parsed.startOverlay = true;
  if (typeof options.logLevel === 'string') parsed.logLevel = parseLogLevel(options.logLevel);
  if (typeof options.passwordStore === 'string') parsed.passwordStore = options.passwordStore;
  if (options.rofi === true) parsed.useRofi = true;
  if (options.startOverlay === true) parsed.autoStartOverlay = true;
  if (options.texthooker === false) parsed.useTexthooker = false;
  if (typeof rootTarget === 'string' && rootTarget) ensureTarget(rootTarget, parsed);
}

export function applyInvocationsToArgs(parsed: Args, invocations: CliInvocations): void {
  if (invocations.doctorTriggered) parsed.doctor = true;
  if (invocations.texthookerTriggered) parsed.texthookerOnly = true;

  if (invocations.jellyfinInvocation) {
    if (invocations.jellyfinInvocation.logLevel) {
      parsed.logLevel = parseLogLevel(invocations.jellyfinInvocation.logLevel);
    }
    if (typeof invocations.jellyfinInvocation.passwordStore === 'string') {
      parsed.passwordStore = invocations.jellyfinInvocation.passwordStore;
    }
    const action = (invocations.jellyfinInvocation.action || '').toLowerCase();
    if (action && !['setup', 'discovery', 'play', 'login', 'logout'].includes(action)) {
      fail(`Unknown jellyfin action: ${invocations.jellyfinInvocation.action}`);
    }
    parsed.jellyfinServer = invocations.jellyfinInvocation.server || '';
    parsed.jellyfinUsername = invocations.jellyfinInvocation.username || '';
    parsed.jellyfinPassword = invocations.jellyfinInvocation.password || '';

    const modeFlags = {
      setup: invocations.jellyfinInvocation.setup || action === 'setup',
      discovery: invocations.jellyfinInvocation.discovery || action === 'discovery',
      play: invocations.jellyfinInvocation.play || action === 'play',
      login: invocations.jellyfinInvocation.login || action === 'login',
      logout: invocations.jellyfinInvocation.logout || action === 'logout',
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

  if (invocations.ytInvocation) {
    if (invocations.ytInvocation.logLevel)
      parsed.logLevel = parseLogLevel(invocations.ytInvocation.logLevel);
    if (invocations.ytInvocation.mode)
      parsed.youtubeSubgenMode = parseYoutubeMode(invocations.ytInvocation.mode);
    if (invocations.ytInvocation.outDir)
      parsed.youtubeSubgenOutDir = invocations.ytInvocation.outDir;
    if (invocations.ytInvocation.keepTemp) parsed.youtubeSubgenKeepTemp = true;
    if (invocations.ytInvocation.whisperBin)
      parsed.whisperBin = invocations.ytInvocation.whisperBin;
    if (invocations.ytInvocation.whisperModel)
      parsed.whisperModel = invocations.ytInvocation.whisperModel;
    if (invocations.ytInvocation.ytSubgenAudioFormat) {
      parsed.youtubeSubgenAudioFormat = invocations.ytInvocation.ytSubgenAudioFormat;
    }
    if (invocations.ytInvocation.target) ensureTarget(invocations.ytInvocation.target, parsed);
  }

  if (invocations.doctorLogLevel) parsed.logLevel = parseLogLevel(invocations.doctorLogLevel);
  if (invocations.texthookerLogLevel)
    parsed.logLevel = parseLogLevel(invocations.texthookerLogLevel);

  if (invocations.configInvocation) {
    if (invocations.configInvocation.logLevel) {
      parsed.logLevel = parseLogLevel(invocations.configInvocation.logLevel);
    }
    const action = (invocations.configInvocation.action || 'path').toLowerCase();
    if (action === 'path') parsed.configPath = true;
    else if (action === 'show') parsed.configShow = true;
    else fail(`Unknown config action: ${invocations.configInvocation.action}`);
  }

  if (invocations.mpvInvocation) {
    if (invocations.mpvInvocation.logLevel) {
      parsed.logLevel = parseLogLevel(invocations.mpvInvocation.logLevel);
    }
    const action = (invocations.mpvInvocation.action || 'status').toLowerCase();
    if (action === 'status') parsed.mpvStatus = true;
    else if (action === 'socket') parsed.mpvSocket = true;
    else if (action === 'idle' || action === 'start') parsed.mpvIdle = true;
    else fail(`Unknown mpv action: ${invocations.mpvInvocation.action}`);
  }

  if (invocations.appInvocation) {
    parsed.appPassthrough = true;
    parsed.appArgs = invocations.appInvocation.appArgs;
  }
}
