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

/**
 * Whether an episode has already been watched, as the stats database records
 * it. Playback marks an episode watched once a session passes the completion
 * threshold, so this is history the app already keeps rather than a second
 * list maintained by the browser.
 */
export interface AnimeBrowserEpisodeWatchState {
  /** The episode's own url, matching the entry it belongs to. */
  episodeUrl: string;
  watched: boolean;
  /** Start of the most recent session, or null when it was never played. */
  lastWatchedMs: number | null;
  sessionCount: number;
}

export interface AnimeBrowserWatchStateRequest {
  sourceId: string;
  animeUrl: string;
  episodeUrls: string[];
}

/**
 * One episode a manual mark applies to. The name and number ride along because
 * marking an episode nobody has played yet has to create its stats row, and
 * that row wants the same series/season/episode fields playback would record.
 */
export interface AnimeBrowserEpisodeMark {
  episodeUrl: string;
  episodeName: string;
  episodeNumber: number | null;
}

export interface AnimeBrowserSetWatchedRequest {
  sourceId: string;
  animeUrl: string;
  animeTitle: string;
  episodes: AnimeBrowserEpisodeMark[];
  watched: boolean;
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
  | 'extracting'
  | 'starting'
  | 'ready'
  | 'failed';

/** Where the running bridge came from, and whether SubMiner can move it forward. */
export interface AnimeBrowserBridgeInstall {
  /**
   * `managed`: downloaded by SubMiner into its own directory, so it can be
   * updated from the browser. `system`: a package-manager install (the AUR
   * `mangatan-extension-server` package) or an `anime.bridgeDir` the user
   * pointed at; SubMiner uses it as found and never writes to it.
   */
  origin: 'managed' | 'system';
  /** Release tag, e.g. `v1.0.6.2`, or null when it cannot be read. */
  version: string | null;
  dir: string;
  /**
   * The newest upstream release with a bundle for this platform, when it is
   * newer than a managed install; null when current, not managed, or not yet
   * checked. Filled in once the bridge is running, since it needs the network.
   */
  updateAvailable: string | null;
}

export interface AnimeBrowserBridgeState {
  stage: AnimeBrowserBridgeStage;
  /** 0-1 while downloading, otherwise null. */
  progress: number | null;
  message: string | null;
  /** Null until the bridge binaries have been located. */
  install: AnimeBrowserBridgeInstall | null;
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
  /** Monotonic Android version code used to decide whether an update exists. */
  versionCode: number;
  nsfw: boolean;
  repoUrl: string;
  /** Where the repository publishes the extension's icon; may 404. */
  iconUrl: string;
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
  /** Read from AndroidManifest.xml; null only for an invalid or unusual APK. */
  versionCode: number | null;
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

/** The Anime Browser episode mpv is currently playing, shared by every browser surface. */
export interface AnimeBrowserPlaybackState {
  sourceId: string;
  animeUrl: string;
  episodeUrl: string;
}

/** One resolved episode waiting in mpv's playlist. */
export type AnimeBrowserQueueEntry = AnimeBrowserPlayRequest;

export interface AnimeBrowserQueueState {
  /** In mpv play order; the first entry starts when the current episode ends. */
  entries: AnimeBrowserQueueEntry[];
  /**
   * Why the last automatic advance failed, or null. Cleared by the next queue
   * change, so it reports the failure the user has not seen yet rather than
   * accumulating a history.
   */
  lastError: string | null;
  /**
   * How many times the queue has started an episode by itself. A counter
   * rather than a flag: it tells a browser window that just repainted whether
   * an advance happened since the state it last saw, including one that
   * started the same episode twice.
   */
  advances: number;
  /** The episode the last advance started, or null before the first one. */
  lastStarted: AnimeBrowserQueueEntry | null;
}

export interface AnimeBrowserAPI {
  getSnapshot: () => Promise<AnimeBrowserSnapshot>;
  ensureBridge: () => Promise<AnimeBrowserBridgeState>;
  /**
   * Download the newest release over a managed install and restart the bridge
   * on it. Rejected for a system install. A playing episode's stream dies with
   * the old bridge.
   */
  updateBridge: () => Promise<AnimeBrowserBridgeState>;
  selectSource: (sourceId: string) => Promise<void>;
  search: (query: string, page?: number) => Promise<AnimeBrowserSearchResult>;
  getPopular: (page?: number) => Promise<AnimeBrowserSearchResult>;
  /** `sourceId` is required after an all-sources search; pass the entry's own. */
  getDetails: (animeUrl: string, sourceId?: string) => Promise<AnimeBrowserDetails>;
  getEpisodes: (animeUrl: string, sourceId?: string) => Promise<AnimeBrowserEpisode[]>;
  /** Watch marks for the listed episodes; empty when stats tracking is off. */
  getWatchState: (
    request: AnimeBrowserWatchStateRequest,
  ) => Promise<AnimeBrowserEpisodeWatchState[]>;
  /** Set or clear the mark by hand; resolves to the state after the write. */
  setWatched: (request: AnimeBrowserSetWatchedRequest) => Promise<AnimeBrowserEpisodeWatchState[]>;
  /** Plays now, replacing whatever mpv is playing. */
  playEpisode: (request: AnimeBrowserPlayRequest) => Promise<AnimeBrowserPlayResult>;
  /** Adds to the end of the queue; queueing an episode twice is a no-op. */
  queueEpisode: (request: AnimeBrowserPlayRequest) => Promise<AnimeBrowserQueueState>;
  dequeueEpisode: (sourceId: string, episodeUrl: string) => Promise<AnimeBrowserQueueState>;
  clearQueue: () => Promise<AnimeBrowserQueueState>;
  getQueue: () => Promise<AnimeBrowserQueueState>;
  getPlaybackState: () => Promise<AnimeBrowserPlaybackState | null>;
  /**
   * Whether mpv has a file open. False when it is idle or not running at all,
   * which is when queueing has no end to wait for.
   */
  isPlaying: () => Promise<boolean>;
  getPreferences: (sourceId: string) => Promise<SourcePreferenceView[]>;
  setPreference: (
    sourceId: string,
    key: string,
    value: string | string[] | boolean,
  ) => Promise<SourcePreferenceView[]>;
  listAvailableExtensions: () => Promise<AvailableExtensionsResult>;
  installExtension: (pkg: string) => Promise<void>;
  updateAllExtensions: () => Promise<number>;
  removeExtension: (pkg: string) => Promise<void>;
  rescanExtensions: () => Promise<void>;
  addRepo: (url: string) => Promise<void>;
  removeRepo: (url: string) => Promise<void>;
  onBridgeState: (listener: (state: AnimeBrowserBridgeState) => void) => () => void;
  onSearchUpdate: (listener: (update: AnimeBrowserSearchUpdate) => void) => () => void;
  /** Pushed whenever the queue changes, including when it advances by itself. */
  onQueueState: (listener: (state: AnimeBrowserQueueState) => void) => () => void;
  /** Pushed when mpv starts an Anime Browser episode or moves to other media. */
  onPlaybackState: (listener: (state: AnimeBrowserPlaybackState | null) => void) => () => void;
}

export type { SourcePreferenceView } from '../anime-bridge/preferences';
