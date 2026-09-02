import type { YoutubeTrackKind } from '../core/services/youtube/kinds';

export type JimakuLanguagePreference = 'ja' | 'en' | 'none';
export type { YoutubeTrackKind };

export interface YoutubeTrackOption {
  id: string;
  language: string;
  sourceLanguage: string;
  kind: YoutubeTrackKind;
  label: string;
  title?: string;
  downloadUrl?: string;
  fileExtension?: string;
}

export interface YoutubePickerOpenPayload {
  sessionId: string;
  url: string;
  tracks: YoutubeTrackOption[];
  defaultPrimaryTrackId: string | null;
  defaultSecondaryTrackId: string | null;
  hasTracks: boolean;
}

export type YoutubePickerResolveRequest =
  | {
      sessionId: string;
      action: 'continue-without-subtitles';
      primaryTrackId: null;
      secondaryTrackId: null;
    }
  | {
      sessionId: string;
      action: 'use-selected';
      primaryTrackId: string | null;
      secondaryTrackId: string | null;
    };

export interface YoutubePickerResolveResult {
  ok: boolean;
  message: string;
}

export interface AnimeConfig {
  /**
   * Directory holding Aniyomi extension `.apk` files. Defaults to
   * `<userData>/anime-extensions` when unset.
   */
  extensionsDir?: string;
  /**
   * Extension repository index URLs (any https `.json` index, such as
   * `https://.../index.min.json`).
   *
   * Ships empty and stays empty unless the user adds one. SubMiner performs no
   * repository discovery and bundles no sources.
   */
  repos?: string[];
  /** Preferred stream label, matched as a substring, e.g. "1080". */
  preferredQuality?: string;
  /**
   * Directory holding an M-Extension-Server bundle (java runtime plus server
   * jar) to run instead of the copy SubMiner downloads. Empty checks the
   * package-manager location, then the managed copy.
   */
  bridgeDir?: string;
}

export interface JimakuConfig {
  apiKey?: string;
  apiKeyCommand?: string;
  apiBaseUrl?: string;
  languagePreference?: JimakuLanguagePreference;
  maxEntryResults?: number;
}

export type AnilistCharacterDictionaryEvictionPolicy = 'disable' | 'delete';
export type AnilistCharacterDictionaryProfileScope = 'all' | 'active';
export type AnilistCharacterDictionaryCollapsibleSectionKey =
  | 'description'
  | 'characterInformation'
  | 'voicedBy';

export interface AnilistCharacterDictionaryCollapsibleSectionsConfig {
  description?: boolean;
  characterInformation?: boolean;
  voicedBy?: boolean;
}

export interface AnilistCharacterDictionaryConfig {
  refreshTtlHours?: number;
  maxLoaded?: number;
  evictionPolicy?: AnilistCharacterDictionaryEvictionPolicy;
  profileScope?: AnilistCharacterDictionaryProfileScope;
  collapsibleSections?: AnilistCharacterDictionaryCollapsibleSectionsConfig;
}

export interface AnilistConfig {
  enabled?: boolean;
  accessToken?: string;
  characterDictionary?: AnilistCharacterDictionaryConfig;
}

export interface YomitanConfig {
  externalProfilePath?: string;
}

export interface JellyfinConfig {
  enabled?: boolean;
  serverUrl?: string;
  recentServers?: string[];
  username?: string;
  defaultLibraryId?: string;
  remoteControlEnabled?: boolean;
  remoteControlAutoConnect?: boolean;
  autoAnnounce?: boolean;
  pullPictures?: boolean;
  iconCacheDir?: string;
  directPlayPreferred?: boolean;
  directPlayContainers?: string[];
  transcodeVideoCodec?: string;
}

export type DiscordPresenceStylePreset = 'default' | 'meme' | 'japanese' | 'minimal';

export interface DiscordPresenceConfig {
  enabled?: boolean;
  presenceStyle?: DiscordPresenceStylePreset;
  updateIntervalMs?: number;
  debounceMs?: number;
}

export interface AiFeatureConfig {
  enabled?: boolean;
  model?: string;
  systemPrompt?: string;
}

export interface AiConfig {
  enabled?: boolean;
  apiKey?: string;
  apiKeyCommand?: string;
  baseUrl?: string;
  model?: string;
  systemPrompt?: string;
  requestTimeoutMs?: number;
}

export type YoutubeMediaCacheMode = 'direct' | 'background';

export interface YoutubeMediaCacheConfig {
  mode?: YoutubeMediaCacheMode;
  maxHeight?: number;
}

export interface YoutubeConfig {
  primarySubLanguages?: string[];
  mediaCache?: YoutubeMediaCacheConfig;
}

export interface YoutubeSubgenConfig {
  whisperBin?: string;
  whisperModel?: string;
  whisperVadModel?: string;
  whisperThreads?: number;
  fixWithAi?: boolean;
  ai?: AiFeatureConfig;
}

export interface StatsConfig {
  toggleKey?: string;
  markWatchedKey?: string;
  serverPort?: number;
  autoStartServer?: boolean;
  autoOpenBrowser?: boolean;
}

export type ImmersionTrackingRetentionMode = 'preset' | 'advanced';
export type ImmersionTrackingRetentionPreset = 'minimal' | 'balanced' | 'deep-history';

export interface ImmersionTrackingConfig {
  enabled?: boolean;
  dbPath?: string;
  batchSize?: number;
  flushIntervalMs?: number;
  queueCap?: number;
  payloadCapBytes?: number;
  maintenanceIntervalMs?: number;
  retentionMode?: ImmersionTrackingRetentionMode;
  retentionPreset?: ImmersionTrackingRetentionPreset;
  retention?: {
    eventsDays?: number;
    telemetryDays?: number;
    sessionsDays?: number;
    dailyRollupsDays?: number;
    monthlyRollupsDays?: number;
    vacuumIntervalDays?: number;
  };
  lifetimeSummaries?: {
    global?: boolean;
    anime?: boolean;
    media?: boolean;
  };
}

export type JimakuConfidence = 'high' | 'medium' | 'low';

export interface JimakuMediaInfo {
  title: string;
  season: number | null;
  episode: number | null;
  confidence: JimakuConfidence;
  filename: string;
  rawTitle: string;
}

export interface JimakuSearchQuery {
  query: string;
}

export interface JimakuEntryFlags {
  anime?: boolean;
  movie?: boolean;
  adult?: boolean;
  external?: boolean;
  unverified?: boolean;
}

export interface JimakuEntry {
  id: number;
  name: string;
  english_name?: string | null;
  japanese_name?: string | null;
  flags?: JimakuEntryFlags;
  last_modified?: string;
}

export interface JimakuFilesQuery {
  entryId: number;
  episode?: number | null;
}

export interface JimakuFileEntry {
  name: string;
  url: string;
  size: number;
  last_modified: string;
}

export interface JimakuDownloadQuery {
  entryId: number;
  url: string;
  name: string;
}

export interface JimakuApiError {
  error: string;
  code?: number;
  retryAfter?: number;
}

export type JimakuApiResponse<T> = { ok: true; data: T } | { ok: false; error: JimakuApiError };

export type JimakuDownloadResult =
  | { ok: true; path: string }
  | { ok: false; error: JimakuApiError };

export interface TsukihimeSearchQuery {
  query: string;
}

export interface TsukihimeEntry {
  id: number;
  title: string;
  timestamp: number | null;
  totalSize: number | null;
  numFiles: number | null;
  sublangs: string[];
}

export interface TsukihimeFilesQuery {
  entryId: number;
}

export interface TsukihimeSubtitleFile {
  attachmentId: number;
  filename: string;
  lang: string;
  trackName: string | null;
  size: number;
  url: string;
  sourceFilename: string;
}

export interface TsukihimeDownloadQuery {
  entryId: number;
  url: string;
  name: string;
  lang?: string;
}

export type TsukihimeApiResponse<T> = JimakuApiResponse<T>;

export type TsukihimeDownloadResult = JimakuDownloadResult;

export interface TsukihimeConfig {
  apiBaseUrl?: string;
  maxSearchResults?: number;
}
