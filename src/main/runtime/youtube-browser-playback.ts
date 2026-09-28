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
export function createYoutubeBrowserPlaybackRuntime(deps: YoutubeBrowserPlaybackDeps) {
  // Only videos sent from the browser get the flow on playlist advance; launcher-started
  // YouTube playback already runs its own flow and must not be doubled.
  const managedVideoIds = new Set<string>();
  // Videos whose flow is still loading; the flow issues its own loadfile, so the path change
  // that follows must not start a second flow for the same video.
  const loadingVideoIds = new Set<string>();
  let lastSeenVideoId: string | null = null;

  const runFlow = async (url: string, videoId: string): Promise<void> => {
    loadingVideoIds.add(videoId);
    try {
      await deps.runPlaybackFlow(url);
    } finally {
      loadingVideoIds.delete(videoId);
    }
  };

  const startFlow = (url: string, videoId: string): void => {
    void runFlow(url, videoId).catch((error: unknown) => {
      const reason = error instanceof Error ? error.message : String(error);
      deps.logWarn(`YouTube browser playback failed for ${url}: ${reason}`);
      deps.notifyFailure(`YouTube playback failed: ${reason}`);
    });
  };

  const openVideo = async (
    request: YoutubeBrowserVideoRequest,
  ): Promise<YoutubeBrowserVideoResult> => {
    const url = toYoutubeWatchUrl(request.url);
    const videoId = extractYoutubeVideoId(url);
    if (!url || !videoId) {
      return { ok: false, message: 'Not a YouTube video link.' };
    }
    managedVideoIds.add(videoId);

    if (request.action === 'queue' && deps.isMpvPlaying()) {
      deps.appendToMpvPlaylist(url);
      return { ok: true, message: 'Queued in mpv' };
    }

    // Starting mpv can take seconds; repeat clicks during that wait must not start a second flow.
    if (loadingVideoIds.has(videoId)) {
      return { ok: true, message: 'Already opening in mpv' };
    }
    loadingVideoIds.add(videoId);
    let mpvReady = false;
    try {
      mpvReady = await deps.ensureMpvReady();
    } finally {
      if (!mpvReady) loadingVideoIds.delete(videoId);
    }
    if (!mpvReady) {
      return { ok: false, message: 'Could not start mpv.' };
    }
    startFlow(url, videoId);
    return { ok: true, message: 'Opening in mpv' };
  };

  const handleMediaPathChange = (mediaPath: string | null | undefined): void => {
    const videoId = extractYoutubeVideoId(mediaPath);
    const previousVideoId = lastSeenVideoId;
    lastSeenVideoId = videoId;
    if (!videoId || videoId === previousVideoId || loadingVideoIds.has(videoId)) return;
    if (!managedVideoIds.has(videoId)) return;
    const url = toYoutubeWatchUrl(mediaPath);
    if (url) startFlow(url, videoId);
  };

  const handleMpvDisconnected = (): void => {
    managedVideoIds.clear();
    lastSeenVideoId = null;
  };

  return { openVideo, handleMediaPathChange, handleMpvDisconnected };
}
