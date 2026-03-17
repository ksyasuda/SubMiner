import fs from 'node:fs';
import path from 'node:path';
import { fail } from '../log.js';
import type { Args, Backend, LauncherYoutubeSubgenConfig, LogLevel } from '../types.js';
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

function parseBackend(value: string): Backend {
  if (value === 'auto' || value === 'hyprland' || value === 'x11' || value === 'macos') {
    return value as Backend;
  }
  fail(`Invalid backend: ${value} (must be auto, hyprland, x11, or macos)`);
}

function parseDictionaryTarget(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    fail('Dictionary target path is required.');
  }
  if (isUrlTarget(trimmed)) {
    fail('Dictionary target must be a local file or directory path, not a URL.');
  }
  const resolved = path.resolve(resolvePathMaybe(trimmed));
  let stat: fs.Stats | null = null;
  try {
    stat = fs.statSync(resolved);
  } catch {
    stat = null;
  }
  if (!stat || (!stat.isFile() && !stat.isDirectory())) {
    fail(`Dictionary target path must be an existing file or directory: ${trimmed}`);
  }
  return resolved;
}

export function createDefaultArgs(launcherConfig: LauncherYoutubeSubgenConfig): Args {
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
    whisperBin: process.env.SUBMINER_WHISPER_BIN || launcherConfig.whisperBin || '',
    whisperModel: process.env.SUBMINER_WHISPER_MODEL || launcherConfig.whisperModel || '',
    whisperVadModel: process.env.SUBMINER_WHISPER_VAD_MODEL || launcherConfig.whisperVadModel || '',
    whisperThreads: (() => {
      const envValue = Number.parseInt(process.env.SUBMINER_WHISPER_THREADS || '', 10);
      if (Number.isInteger(envValue) && envValue > 0) return envValue;
      return launcherConfig.whisperThreads || 4;
    })(),
    youtubeSubgenOutDir: process.env.SUBMINER_YT_SUBGEN_OUT_DIR || DEFAULT_YOUTUBE_SUBGEN_OUT_DIR,
    youtubeSubgenAudioFormat: process.env.SUBMINER_YT_SUBGEN_AUDIO_FORMAT || 'm4a',
    youtubeSubgenKeepTemp: process.env.SUBMINER_YT_SUBGEN_KEEP_TEMP === '1',
    youtubeFixWithAi: launcherConfig.fixWithAi === true,
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
    dictionary: false,
    stats: false,
    statsCleanup: false,
    statsCleanupVocab: false,
    statsCleanupLifetime: false,
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
    aiConfig: {
      enabled: launcherConfig.ai?.enabled,
      apiKey: launcherConfig.ai?.apiKey,
      apiKeyCommand: launcherConfig.ai?.apiKeyCommand,
      baseUrl: launcherConfig.ai?.baseUrl,
      model: launcherConfig.ai?.model,
      systemPrompt: launcherConfig.ai?.systemPrompt,
      requestTimeoutMs: launcherConfig.ai?.requestTimeoutMs,
    },
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
  if (invocations.dictionaryTriggered) parsed.dictionary = true;
  if (invocations.statsTriggered) parsed.stats = true;
  if (invocations.statsCleanup) parsed.statsCleanup = true;
  if (invocations.statsCleanupVocab) parsed.statsCleanupVocab = true;
  if (invocations.statsCleanupLifetime) parsed.statsCleanupLifetime = true;
  if (invocations.dictionaryTarget) {
    parsed.dictionaryTarget = parseDictionaryTarget(invocations.dictionaryTarget);
  }
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
    if (invocations.ytInvocation.outDir)
      parsed.youtubeSubgenOutDir = invocations.ytInvocation.outDir;
    if (invocations.ytInvocation.keepTemp) parsed.youtubeSubgenKeepTemp = true;
    if (invocations.ytInvocation.whisperBin)
      parsed.whisperBin = invocations.ytInvocation.whisperBin;
    if (invocations.ytInvocation.whisperModel)
      parsed.whisperModel = invocations.ytInvocation.whisperModel;
    if (invocations.ytInvocation.whisperVadModel)
      parsed.whisperVadModel = invocations.ytInvocation.whisperVadModel;
    if (invocations.ytInvocation.whisperThreads)
      parsed.whisperThreads = invocations.ytInvocation.whisperThreads;
    if (invocations.ytInvocation.ytSubgenAudioFormat) {
      parsed.youtubeSubgenAudioFormat = invocations.ytInvocation.ytSubgenAudioFormat;
    }
    if (invocations.ytInvocation.target) ensureTarget(invocations.ytInvocation.target, parsed);
  }

  if (invocations.dictionaryLogLevel) {
    parsed.logLevel = parseLogLevel(invocations.dictionaryLogLevel);
  }
  if (invocations.statsLogLevel) {
    parsed.logLevel = parseLogLevel(invocations.statsLogLevel);
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
