import { extractYoutubeVideoId, toYoutubeWatchUrl } from './youtube-playback';

export type YoutubeBrowserVideoAction = 'play' | 'queue';

export type YoutubeBrowserVideoRequest = {
  action: YoutubeBrowserVideoAction;
  url: string;
};

export type YoutubeBrowserVideoResult = {
  ok: boolean;
  message: string;
};

export type YoutubeBrowserPlaybackDeps = {
  /** True when mpv is connected and has a file loaded, so appending queues behind it. */
  isMpvPlaying: () => boolean;
  /** Connect to (or launch) mpv so the playback flow has a player to drive. */
  ensureMpvReady: () => Promise<boolean>;
  /** SubMiner's YouTube playback flow: load the video, pick subtitles, release the play gate. */
  runPlaybackFlow: (url: string) => Promise<void>;
  appendToMpvPlaylist: (url: string) => void;
  notifyFailure: (message: string) => void;
  logWarn: (message: string) => void;
  /** A play request ended without loading its video, after queued requests behind it were released. */
  onStartupFailed: () => void;
};

export function parseYoutubeBrowserVideoRequest(
  payload: unknown,
): YoutubeBrowserVideoRequest | null {
  if (typeof payload !== 'object' || payload === null) return null;
  const { action, url } = payload as Record<string, unknown>;
  if ((action !== 'play' && action !== 'queue') || typeof url !== 'string') return null;
  return { action, url };
}

/**
 * Hands videos picked in the YouTube browser window to mpv. Queued videos land in mpv's own
 * playlist; when mpv advances onto one, `handleMediaPathChange` runs the playback flow for it so
 * queued videos get the same subtitle setup as videos played directly.
 */
type PlaybackStartup = {
  videoId: string;
  /** Resolves true once mpv has the video loaded, false if mpv or the flow failed first. */
  loaded: Promise<boolean>;
  settle: (loaded: boolean) => void;
};

export function createYoutubeBrowserPlaybackRuntime(deps: YoutubeBrowserPlaybackDeps) {
  // Videos queued from the browser that mpv has not reached yet. Only these get the flow on
  // playlist advance; launcher-started YouTube playback already runs its own flow.
  const queuedVideoIds = new Set<string>();
  // Videos whose flow is still loading; the flow issues its own loadfile, so the path change
  // that follows must not start a second flow for the same video.
  const loadingVideoIds = new Set<string>();
  // The latest play request while mpv starts and loads it. Queue requests arriving meanwhile wait
  // on it so they append behind that video instead of starting a second flow that replaces it.
  let startup: PlaybackStartup | null = null;
  let lastSeenVideoId: string | null = null;

  const beginStartup = (videoId: string): PlaybackStartup => {
    let resolveLoaded: (loaded: boolean) => void = () => {};
    let settled = false;
    const entry: PlaybackStartup = {
      videoId,
      loaded: new Promise<boolean>((resolve) => {
        resolveLoaded = resolve;
      }),
      settle: (loaded) => {
        if (settled) return;
        settled = true;
        if (startup === entry) startup = null;
        resolveLoaded(loaded);
        // Registered after the queued requests' handlers, so one that plays instead has already
        // begun its own startup by the time this runs.
        if (!loaded) void entry.loaded.then(() => deps.onStartupFailed());
      },
    };
    startup = entry;
    return entry;
  };

  /** Runs the playback flow; resolves false (after reporting) when it fails. */
  const startFlow = async (url: string, videoId: string): Promise<boolean> => {
    loadingVideoIds.add(videoId);
    try {
      await deps.runPlaybackFlow(url);
      return true;
    } catch (error: unknown) {
      const reason = error instanceof Error ? error.message : String(error);
      deps.logWarn(`YouTube browser playback failed for ${url}: ${reason}`);
      deps.notifyFailure(`YouTube playback failed: ${reason}`);
      return false;
    } finally {
      loadingVideoIds.delete(videoId);
    }
  };

  const appendToMpv = (url: string, videoId: string): YoutubeBrowserVideoResult => {
    queuedVideoIds.add(videoId);
    deps.appendToMpvPlaylist(url);
    return { ok: true, message: 'Queued in mpv' };
  };

  const queueBehindStartup = (
    pending: PlaybackStartup,
    request: YoutubeBrowserVideoRequest,
    url: string,
    videoId: string,
  ): void => {
    // Chained in click order. If the earlier video never loaded, this one plays instead.
    void pending.loaded
      .then((loaded) => (loaded ? appendToMpv(url, videoId) : openVideo(request)))
      .then(
        (result) => {
          if (!result.ok) deps.notifyFailure(`YouTube queue failed: ${result.message}`);
        },
        (error: unknown) => {
          const reason = error instanceof Error ? error.message : String(error);
          deps.logWarn(`YouTube browser queue failed for ${url}: ${reason}`);
          deps.notifyFailure(`YouTube queue failed: ${reason}`);
        },
      );
  };

  const openVideo = async (
    request: YoutubeBrowserVideoRequest,
  ): Promise<YoutubeBrowserVideoResult> => {
    const url = toYoutubeWatchUrl(request.url);
    const videoId = extractYoutubeVideoId(url);
    if (!url || !videoId) {
      return { ok: false, message: 'Not a YouTube video link.' };
    }

    if (request.action === 'queue') {
      if (startup) {
        queueBehindStartup(startup, request, url, videoId);
        return { ok: true, message: 'Queued in mpv' };
      }
      if (deps.isMpvPlaying()) return appendToMpv(url, videoId);
    }

    // Starting mpv can take seconds; repeat clicks during that wait must not start a second flow.
    if (loadingVideoIds.has(videoId)) {
      return { ok: true, message: 'Already opening in mpv' };
    }
    loadingVideoIds.add(videoId);
    const entry = beginStartup(videoId);
    let mpvReady = false;
    try {
      mpvReady = await deps.ensureMpvReady();
    } finally {
      if (!mpvReady) {
        loadingVideoIds.delete(videoId);
        entry.settle(false);
      }
    }
    if (!mpvReady) {
      return { ok: false, message: 'Could not start mpv.' };
    }
    void startFlow(url, videoId).then(entry.settle);
    return { ok: true, message: 'Opening in mpv' };
  };

  const handleMediaPathChange = (mediaPath: string | null | undefined): void => {
    const videoId = extractYoutubeVideoId(mediaPath);
    const previousVideoId = lastSeenVideoId;
    lastSeenVideoId = videoId;
    if (videoId && startup?.videoId === videoId) startup.settle(true);
    if (!videoId || videoId === previousVideoId || loadingVideoIds.has(videoId)) return;
    if (!queuedVideoIds.delete(videoId)) return;
    const url = toYoutubeWatchUrl(mediaPath);
    if (url) void startFlow(url, videoId);
  };

  const handleMpvDisconnected = (): void => {
    queuedVideoIds.clear();
    lastSeenVideoId = null;
  };

  /** True while a browser play request is still starting mpv or loading its video. */
  const isStartingPlayback = (): boolean => startup !== null;

  return { openVideo, handleMediaPathChange, handleMpvDisconnected, isStartingPlayback };
}
