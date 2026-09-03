import type { AnimeStreamMetadata } from '../../anime-bridge/episode-metadata';
import type {
  StreamPlaybackMetadataInput,
  StreamWatchState,
} from '../../core/services/immersion-tracker-service';

/** What the stats store needs to record a mark against one episode. */
export type StreamWatchMark = StreamPlaybackMetadataInput;
import type { PlaybackEndFileEvent } from '../../anime-bridge/playback-outcome';
import type { SubtitleCacheIo } from '../../anime-bridge/subtitle-cache';
import { startSidecar } from '../../anime-bridge/sidecar-process';
import { startStreamStripProxy } from '../../anime-bridge/stream-strip-proxy';
import type {
  AnimeBrowserBridgeInstall,
  AnimeBrowserBridgeState,
  AnimeBrowserQueueState,
  AnimeBrowserSearchUpdate,
} from '../../types/anime-browser';
import type { BridgeInstall, InstallProgress, StagedBridgeUpdate } from './anime-bridge-installer';

export interface AnimeBrowserRuntimeDeps {
  /** Where user-supplied Aniyomi extension APKs live. Read lazily so config edits apply. */
  extensionsDir: () => string;
  /** Configured repository index URLs. Empty unless the user added one. */
  repos: () => string[];
  /** Persists the repository list. Config stays the source of truth. */
  setRepos: (repos: string[]) => void;
  /** JSON file holding each source's saved preference values. */
  preferencesFile: string;
  /** Locates (or first downloads) the bridge and says where it came from. */
  ensureBinaries: (onProgress: (progress: InstallProgress) => void) => Promise<BridgeInstall>;
  /**
   * The newest upstream release a managed install could move to, or null.
   * Asked once the bridge is running; a rejection is logged, not shown.
   */
  checkBridgeUpdate: (install: AnimeBrowserBridgeInstall) => Promise<string | null>;
  /**
   * Downloads the newest release beside the managed install; the returned
   * `commit` swaps it in once the old bridge has stopped.
   */
  stageBridgeUpdate: (
    onProgress: (progress: InstallProgress) => void,
  ) => Promise<StagedBridgeUpdate>;
  /** Sends mpv an IPC command; same transport the Jellyfin path uses. */
  sendMpvCommand: (command: Array<string | number>) => void;
  /** Brings mpv up if it is not already connected. Resolves false on failure. */
  ensureMpvConnected: () => Promise<boolean>;
  /** Subscribe to mpv end-file events so playback startup can be confirmed. */
  onPlaybackEndFile?: (listener: (event: PlaybackEndFileEvent) => void) => () => void;
  /** Subscribe to the active mpv path for real-playlist anime queue advances. */
  onPlaybackPathChange?: (listener: (path: string) => void) => () => void;
  /** One-shot mpv property read; rejects while the property is unavailable. */
  readMpvProperty?: (name: string) => Promise<unknown>;
  showMpvOsd?: (message: string) => void;
  showVisibleOverlay?: () => void;
  /** Publishes stream identity before loadfile starts the stats session. */
  onPlaybackMetadata?: (metadata: AnimeStreamMetadata) => void;
  /** Registers a queued stream before mpv can navigate to its playlist entry. */
  onPreparedPlaybackMetadata?: (metadata: AnimeStreamMetadata) => void;
  /**
   * Watch state for the given stats paths, from the same store playback writes
   * to. Absent (or resolving empty) when stats tracking is disabled, which the
   * browser shows as "no watch history" rather than as an error.
   */
  getWatchState?: (statsPaths: string[]) => Promise<Map<string, StreamWatchState>>;
  /**
   * Sets or clears the watch mark by hand. Marking creates the stats row for an
   * episode nobody has played yet, which is what makes catching up on a series
   * watched elsewhere possible.
   */
  setWatchState?: (episodes: StreamWatchMark[], watched: boolean) => Promise<number>;
  /** Lets tests drive the pause between loadfile and track attachment. */
  wait?: (ms: number) => Promise<void>;
  /** Overrides the filesystem/network the subtitle cache uses. Tests only. */
  subtitleCacheIo?: SubtitleCacheIo;
  onBridgeState: (state: AnimeBrowserBridgeState) => void;
  /** Pushes the play queue to the browser window, advances included. */
  onQueueState?: (state: AnimeBrowserQueueState) => void;
  /** Streams per-source progress while a search invoke is pending. */
  onSearchUpdate?: (update: AnimeBrowserSearchUpdate, sessionId: string) => void;
  preferredQuality?: () => string | undefined;
  /**
   * Source id (or `all`) a new browser session starts on. Falls back to the
   * first installed source when unset or no longer installed.
   */
  defaultSourceId?: () => string | undefined;
  /** Persists the default; absent when the host has no config to write to. */
  setDefaultSourceId?: (sourceId: string) => void;
  log: (message: string) => void;
  /** Overrides process startup in focused runtime tests. */
  startSidecar?: typeof startSidecar;
  /** Overrides proxy startup in focused runtime tests. */
  startStreamStripProxy?: typeof startStreamStripProxy;
}

export type AnimeBrowserPlaybackDeps = Pick<
  AnimeBrowserRuntimeDeps,
  | 'sendMpvCommand'
  | 'ensureMpvConnected'
  | 'onPlaybackEndFile'
  | 'readMpvProperty'
  | 'showMpvOsd'
  | 'showVisibleOverlay'
  | 'onPlaybackMetadata'
  | 'onPreparedPlaybackMetadata'
  | 'wait'
  | 'subtitleCacheIo'
  | 'preferredQuality'
  | 'log'
>;
