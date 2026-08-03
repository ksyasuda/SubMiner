import type { AnimeStreamMetadata } from '../../anime-bridge/episode-metadata';
import type {
  StreamPlaybackMetadataInput,
  StreamWatchState,
} from '../../core/services/immersion-tracker-service';

/** What the stats store needs to record a mark against one episode. */
export type StreamWatchMark = StreamPlaybackMetadataInput;
import type { PlaybackEndFileEvent } from '../../anime-bridge/playback-outcome';
import type { SubtitleCacheIo } from '../../anime-bridge/subtitle-cache';
import type { BundleBinaries } from '../../anime-bridge/sidecar-bundle';
import { startSidecar } from '../../anime-bridge/sidecar-process';
import { startStreamStripProxy } from '../../anime-bridge/stream-strip-proxy';
import type { AnimeBrowserBridgeState, AnimeBrowserSearchUpdate } from '../../types/anime-browser';
import type { InstallProgress } from './anime-bridge-installer';

export interface AnimeBrowserRuntimeDeps {
  /** Where user-supplied Aniyomi extension APKs live. Read lazily so config edits apply. */
  extensionsDir: () => string;
  /** Configured repository index URLs. Empty unless the user added one. */
  repos: () => string[];
  /** Persists the repository list. Config stays the source of truth. */
  setRepos: (repos: string[]) => void;
  /** JSON file holding each source's saved preference values. */
  preferencesFile: string;
  ensureBinaries: (onProgress: (progress: InstallProgress) => void) => Promise<BundleBinaries>;
  /** Sends mpv an IPC command; same transport the Jellyfin path uses. */
  sendMpvCommand: (command: Array<string | number>) => void;
  /** Brings mpv up if it is not already connected. Resolves false on failure. */
  ensureMpvConnected: () => Promise<boolean>;
  /** Subscribe to mpv end-file events so playback startup can be confirmed. */
  onPlaybackEndFile?: (listener: (event: PlaybackEndFileEvent) => void) => () => void;
  /** One-shot mpv property read; rejects while the property is unavailable. */
  readMpvProperty?: (name: string) => Promise<unknown>;
  showMpvOsd?: (message: string) => void;
  showVisibleOverlay?: () => void;
  /** Publishes stream identity before loadfile starts the stats session. */
  onPlaybackMetadata?: (metadata: AnimeStreamMetadata) => void;
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
  /** Streams per-source progress while a search invoke is pending. */
  onSearchUpdate?: (update: AnimeBrowserSearchUpdate) => void;
  preferredQuality?: () => string | undefined;
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
  | 'wait'
  | 'subtitleCacheIo'
  | 'preferredQuality'
  | 'log'
>;
