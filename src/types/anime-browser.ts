import type { AnimeStatus } from '../anime-bridge/media-url';
import type { SourcePreferenceView } from '../anime-bridge/preferences';

/**
 * Stands in for a source id when every installed source should answer. Not a
 * real source: `getPreferences` and the per-anime calls still need a single one.
 */
export const ALL_SOURCES_ID = '__all__';

/** An installed anime extension the browser can search. */
export interface AnimeBrowserSource {
  id: string;
  name: string;
  lang: string;
  /** Package name, e.g. `eu.kanade.tachiyomi.animeextension.all.jellyfin`. */
  pkg: string;
}

export interface AnimeBrowserEntry {
  /** Source-relative url; the handle for every later call. */
  url: string;
  title: string;
  thumbnailUrl: string | null;
  /**
   * Which source produced this entry. Carried on the entry rather than read
   * from the current selection, so an all-sources result stays usable.
   */
  sourceId: string;
  sourceName: string;
}

export interface AnimeBrowserDetails extends AnimeBrowserEntry {
  description: string | null;
  author: string | null;
  genres: string[];
  status: AnimeStatus;
}

export interface AnimeBrowserEpisode {
  url: string;
  name: string;
  /** Extension-reported episode number; may be fractional for specials. */
  number: number | null;
  /** Epoch milliseconds, or null when the source reports no date. */
  uploadedAt: number | null;
  scanlator: string | null;
}

/** One source that errored while the others answered. */
export interface SourceSearchFailure {
  sourceId: string;
  sourceName: string;
  error: string;
}

export interface AnimeBrowserSearchResult {
  entries: AnimeBrowserEntry[];
  hasNextPage: boolean;
  /**
   * Sources that failed during an all-sources search. A single-source search
   * rejects instead, so this is empty there.
   */
  failures: SourceSearchFailure[];
}

/**
 * Incremental progress of one search, pushed while the search invoke is still
 * pending so a fast source is visible before a slow one answers.
 *
 * `token` orders searches: the renderer keeps the highest `start` token it has
 * seen and drops events from any other search, so a stale search that is still
 * resolving cannot paint over the one the user just typed.
 */
export type AnimeBrowserSearchUpdate =
  | { kind: 'start'; token: number; sourceCount: number }
  | {
      kind: 'result';
      token: number;
      sourceId: string;
      sourceName: string;
      entries: AnimeBrowserEntry[];
    }
  | { kind: 'failure'; token: number; failure: SourceSearchFailure }
  | { kind: 'done'; token: number };

/** Progress of the one long-running operation: bringing the bridge up. */
export type AnimeBrowserBridgeStage =
  | 'idle'
  | 'locating'
  | 'downloading'
  | 'verifying'
  | 'extracting'
  | 'starting'
  | 'ready'
  | 'failed';

export interface AnimeBrowserBridgeState {
  stage: AnimeBrowserBridgeStage;
  /** 0-1 while downloading, otherwise null. */
  progress: number | null;
  message: string | null;
}

/** An extension APK that failed to load, surfaced instead of silently vanishing. */
export interface ExtensionLoadFailure {
  /** Package (file) name of the APK that failed. */
  pkg: string;
  error: string;
}

/** An extension offered by a configured repository. */
export interface AvailableExtension {
  pkg: string;
  name: string;
  lang: string;
  version: string;
  nsfw: boolean;
  repoUrl: string;
  sourceNames: string[];
  installed: boolean;
}

export interface RepoFailure {
  repoUrl: string;
  error: string;
}

export interface AvailableExtensionsResult {
  extensions: AvailableExtension[];
  failures: RepoFailure[];
}

/**
 * An extension present in the extensions directory. Listed from disk rather
 * than from a repository, so an APK dropped in by hand — or one whose
 * repository has since been removed — can still be seen and removed.
 */
export interface InstalledExtensionView {
  pkg: string;
  /** The sources it provides, or the file name when it provided none. */
  name: string;
  /** Languages its sources cover, deduplicated. */
  langs: string[];
  /** How many sources it provides; 0 when it failed to load. */
  sourceCount: number;
  /** Why it failed to load, or null when it loaded. */
  error: string | null;
}

export interface AnimeBrowserSnapshot {
  bridge: AnimeBrowserBridgeState;
  sources: AnimeBrowserSource[];
  selectedSourceId: string | null;
  loadFailures: ExtensionLoadFailure[];
  /** Every extension on disk, whether or not it loaded. */
  installed: InstalledExtensionView[];
  /** Where APKs are read from, shown so the user knows where to drop files. */
  extensionsDir: string;
  /** Configured repository index URLs. Empty until the user adds one. */
  repos: string[];
}

export interface AnimeBrowserPlayRequest {
  sourceId: string;
  animeUrl: string;
  animeTitle: string;
  episodeUrl: string;
  episodeName: string;
  /**
   * The episode's number as the source reported it. Carried rather than
   * re-parsed out of `episodeName`, which is free-form and often lacks one.
   */
  episodeNumber: number | null;
}

export interface AnimeBrowserPlayResult {
  ok: boolean;
  /** Populated when ok is false, phrased for display. */
  error: string | null;
  quality: string | null;
}

export interface AnimeBrowserAPI {
  getSnapshot: () => Promise<AnimeBrowserSnapshot>;
  ensureBridge: () => Promise<AnimeBrowserBridgeState>;
  selectSource: (sourceId: string) => Promise<void>;
  search: (query: string, page?: number) => Promise<AnimeBrowserSearchResult>;
  getPopular: (page?: number) => Promise<AnimeBrowserSearchResult>;
  /** `sourceId` is required after an all-sources search; pass the entry's own. */
  getDetails: (animeUrl: string, sourceId?: string) => Promise<AnimeBrowserDetails>;
  getEpisodes: (animeUrl: string, sourceId?: string) => Promise<AnimeBrowserEpisode[]>;
  playEpisode: (request: AnimeBrowserPlayRequest) => Promise<AnimeBrowserPlayResult>;
  getPreferences: (sourceId: string) => Promise<SourcePreferenceView[]>;
  setPreference: (
    sourceId: string,
    key: string,
    value: string | string[] | boolean,
  ) => Promise<SourcePreferenceView[]>;
  listAvailableExtensions: () => Promise<AvailableExtensionsResult>;
  installExtension: (pkg: string) => Promise<void>;
  removeExtension: (pkg: string) => Promise<void>;
  rescanExtensions: () => Promise<void>;
  addRepo: (url: string) => Promise<void>;
  removeRepo: (url: string) => Promise<void>;
  onBridgeState: (listener: (state: AnimeBrowserBridgeState) => void) => () => void;
  onSearchUpdate: (listener: (update: AnimeBrowserSearchUpdate) => void) => () => void;
}

export type { SourcePreferenceView } from '../anime-bridge/preferences';
