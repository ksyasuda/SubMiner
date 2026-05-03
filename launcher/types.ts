import path from 'node:path';
import os from 'node:os';
import type { MpvLaunchMode } from '../src/types/config.js';
import { resolveDefaultLogFilePath } from '../src/shared/log-files.js';
export { VIDEO_EXTENSIONS } from '../src/shared/video-extensions.js';

export const ROFI_THEME_FILE = 'subminer.rasi';
export function getDefaultSocketPath(platform: NodeJS.Platform = process.platform): string {
  if (platform === 'win32') {
    return '\\\\.\\pipe\\subminer-socket';
  }
  return '/tmp/subminer-socket';
}

export const DEFAULT_SOCKET_PATH = getDefaultSocketPath();
export const DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS = ['ja', 'jpn'];
export const DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS = ['en', 'eng'];
export const YOUTUBE_SUB_EXTENSIONS = new Set(['.srt', '.vtt', '.ass']);
export const YOUTUBE_AUDIO_EXTENSIONS = new Set([
  '.m4a',
  '.mp3',
  '.webm',
  '.opus',
  '.wav',
  '.aac',
  '.flac',
]);
export const DEFAULT_YOUTUBE_SUBGEN_OUT_DIR = path.join(
  os.homedir(),
  '.cache',
  'subminer',
  'youtube-subs',
);
export function getDefaultLauncherLogFile(options?: {
  platform?: NodeJS.Platform;
  homeDir?: string;
  appDataDir?: string;
}): string {
  return resolveDefaultLogFilePath('launcher', {
    platform: options?.platform ?? process.platform,
    homeDir: options?.homeDir ?? os.homedir(),
    appDataDir: options?.appDataDir,
  });
}

export function getDefaultMpvLogFile(options?: {
  platform?: NodeJS.Platform;
  homeDir?: string;
  appDataDir?: string;
}): string {
  return resolveDefaultLogFilePath('mpv', {
    platform: options?.platform ?? process.platform,
    homeDir: options?.homeDir ?? os.homedir(),
    appDataDir: options?.appDataDir,
  });
}

export const DEFAULT_MPV_LOG_FILE = getDefaultMpvLogFile();
export const DEFAULT_YOUTUBE_YTDL_FORMAT = 'bestvideo*+bestaudio/best';
export const DEFAULT_JIMAKU_API_BASE_URL = 'https://jimaku.cc';
export const DEFAULT_MPV_SUBMINER_ARGS = [
  '--sub-auto=fuzzy',
  '--sub-file-paths=.;subs;subtitles',
  '--sid=auto',
  '--secondary-sid=auto',
  '--secondary-sub-visibility=no',
  '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
  '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
] as const;

export type LogLevel = 'debug' | 'info' | 'warn' | 'error';
export type Backend = 'auto' | 'hyprland' | 'sway' | 'x11' | 'macos' | 'windows';
export type JimakuLanguagePreference = 'ja' | 'en' | 'none';

export interface LauncherAiConfig {
  enabled?: boolean;
  apiKey?: string;
  apiKeyCommand?: string;
  baseUrl?: string;
  model?: string;
  systemPrompt?: string;
  requestTimeoutMs?: number;
}

export interface Args {
  backend: Backend;
  directory: string;
  recursive: boolean;
  profile: string;
  startOverlay: boolean;
  youtubeMode?: 'download' | 'generate';
  whisperBin: string;
  whisperModel: string;
  whisperVadModel: string;
  whisperThreads: number;
  youtubeSubgenOutDir: string;
  youtubeSubgenAudioFormat: string;
  youtubeSubgenKeepTemp: boolean;
  youtubeFixWithAi: boolean;
  youtubePrimarySubLangs: string[];
  youtubeSecondarySubLangs: string[];
  youtubeAudioLangs: string[];
  youtubeWhisperSourceLanguage: string;
  aiConfig: LauncherAiConfig;
  useTexthooker: boolean;
  autoStartOverlay: boolean;
  texthookerOnly: boolean;
  texthookerOpenBrowser: boolean;
  useRofi: boolean;
  logLevel: LogLevel;
  passwordStore: string;
  target: string;
  targetKind: '' | 'file' | 'url';
  jimakuApiKey: string;
  jimakuApiKeyCommand: string;
  jimakuApiBaseUrl: string;
  jimakuLanguagePreference: JimakuLanguagePreference;
  jimakuMaxEntryResults: number;
  jellyfin: boolean;
  jellyfinLogin: boolean;
  jellyfinLogout: boolean;
  jellyfinPlay: boolean;
  jellyfinDiscovery: boolean;
  dictionary: boolean;
  dictionaryCandidates: boolean;
  dictionarySelect: boolean;
  dictionaryAnilistId?: number;
  stats: boolean;
  statsBackground?: boolean;
  statsStop?: boolean;
  statsCleanup?: boolean;
  statsCleanupVocab?: boolean;
  statsCleanupLifetime?: boolean;
  dictionaryTarget?: string;
  doctor: boolean;
  doctorRefreshKnownWords: boolean;
  configPath: boolean;
  configShow: boolean;
  mpvIdle: boolean;
  mpvSocket: boolean;
  mpvStatus: boolean;
  mpvArgs: string;
  appPassthrough: boolean;
  appArgs: string[];
  jellyfinServer: string;
  jellyfinUsername: string;
  jellyfinPassword: string;
  launchMode: MpvLaunchMode;
}

export interface LauncherYoutubeSubgenConfig {
  whisperBin?: string;
  whisperModel?: string;
  whisperVadModel?: string;
  whisperThreads?: number;
  fixWithAi?: boolean;
  ai?: LauncherAiConfig;
  primarySubLanguages?: string[];
  secondarySubLanguages?: string[];
  jimakuApiKey?: string;
  jimakuApiKeyCommand?: string;
  jimakuApiBaseUrl?: string;
  jimakuLanguagePreference?: JimakuLanguagePreference;
  jimakuMaxEntryResults?: number;
}

export interface LauncherJellyfinConfig {
  enabled?: boolean;
  serverUrl?: string;
  username?: string;
  defaultLibraryId?: string;
  pullPictures?: boolean;
  iconCacheDir?: string;
}

export interface LauncherMpvConfig {
  launchMode?: MpvLaunchMode;
}

export interface PluginRuntimeConfig {
  socketPath: string;
  autoStart: boolean;
  autoStartVisibleOverlay: boolean;
  autoStartPauseUntilReady: boolean;
}

export interface CommandExecOptions {
  allowFailure?: boolean;
  captureStdout?: boolean;
  logLevel?: LogLevel;
  commandLabel?: string;
  streamOutput?: boolean;
  env?: NodeJS.ProcessEnv;
}

export interface CommandExecResult {
  code: number;
  stdout: string;
  stderr: string;
}

export interface SubtitleCandidate {
  path: string;
  lang: 'primary' | 'secondary';
  ext: string;
  size: number;
  source: 'manual' | 'whisper' | 'whisper-fixed' | 'whisper-translate' | 'whisper-translate-fixed';
}

export interface YoutubeSubgenOutputs {
  basename: string;
  primaryPath?: string;
  secondaryPath?: string;
  primaryNative?: boolean;
  secondaryNative?: boolean;
}

export interface MpvTrack {
  type?: string;
  id?: number;
  lang?: string;
  title?: string;
}

export interface JellyfinSessionConfig {
  serverUrl: string;
  accessToken: string;
  userId: string;
  defaultLibraryId: string;
  pullPictures: boolean;
  iconCacheDir: string;
}

export interface JellyfinLibraryEntry {
  id: string;
  name: string;
  kind: string;
}

export interface JellyfinItemEntry {
  id: string;
  name: string;
  type: string;
  display: string;
}

export interface JellyfinGroupEntry {
  id: string;
  name: string;
  type: string;
  display: string;
}
