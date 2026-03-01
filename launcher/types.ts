import path from 'node:path';
import os from 'node:os';

export const VIDEO_EXTENSIONS = new Set([
  'mkv',
  'mp4',
  'avi',
  'webm',
  'mov',
  'flv',
  'wmv',
  'm4v',
  'ts',
  'm2ts',
]);

export const ROFI_THEME_FILE = 'subminer.rasi';
export const DEFAULT_SOCKET_PATH = '/tmp/subminer-socket';
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
export const DEFAULT_MPV_LOG_FILE = path.join(
  os.homedir(),
  '.config',
  'SubMiner',
  'logs',
  `SubMiner-${new Date().toISOString().slice(0, 10)}.log`,
);
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
export type YoutubeSubgenMode = 'automatic' | 'preprocess' | 'off';
export type Backend = 'auto' | 'hyprland' | 'x11' | 'macos';
export type JimakuLanguagePreference = 'ja' | 'en' | 'none';

export interface Args {
  backend: Backend;
  directory: string;
  recursive: boolean;
  profile: string;
  startOverlay: boolean;
  youtubeSubgenMode: YoutubeSubgenMode;
  whisperBin: string;
  whisperModel: string;
  youtubeSubgenOutDir: string;
  youtubeSubgenAudioFormat: string;
  youtubeSubgenKeepTemp: boolean;
  youtubePrimarySubLangs: string[];
  youtubeSecondarySubLangs: string[];
  youtubeAudioLangs: string[];
  youtubeWhisperSourceLanguage: string;
  useTexthooker: boolean;
  autoStartOverlay: boolean;
  texthookerOnly: boolean;
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
  doctor: boolean;
  configPath: boolean;
  configShow: boolean;
  mpvIdle: boolean;
  mpvSocket: boolean;
  mpvStatus: boolean;
  appPassthrough: boolean;
  appArgs: string[];
  jellyfinServer: string;
  jellyfinUsername: string;
  jellyfinPassword: string;
}

export interface LauncherYoutubeSubgenConfig {
  mode?: YoutubeSubgenMode;
  whisperBin?: string;
  whisperModel?: string;
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
  source: 'manual' | 'auto' | 'whisper' | 'whisper-translate';
}

export interface YoutubeSubgenOutputs {
  basename: string;
  primaryPath?: string;
  secondaryPath?: string;
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
