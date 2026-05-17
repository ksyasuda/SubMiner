type MpvBindingEventName =
  | 'connection-change'
  | 'subtitle-change'
  | 'subtitle-ass-change'
  | 'secondary-subtitle-change'
  | 'subtitle-track-change'
  | 'subtitle-track-list-change'
  | 'subtitle-timing'
  | 'media-path-change'
  | 'media-title-change'
  | 'time-pos-change'
  | 'duration-change'
  | 'pause-change'
  | 'fullscreen-change'
  | 'subtitle-metrics-change'
  | 'secondary-subtitle-visibility';

type MpvEventClient = {
  on: <K extends MpvBindingEventName>(event: K, handler: (payload: any) => void) => void;
};

type AnilistPostWatchRunOptions = {
  watchedSeconds?: number;
};

export function createHandleMpvConnectionChangeHandler(deps: {
  reportJellyfinRemoteStopped: () => void;
  refreshDiscordPresence: () => void;
  syncOverlayMpvSubtitleSuppression: () => void;
  hasInitialPlaybackQuitOnDisconnectArg: () => boolean;
  isOverlayRuntimeInitialized: () => boolean;
  shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => boolean;
  isQuitOnDisconnectArmed: () => boolean;
  scheduleQuitCheck: (callback: () => void) => void;
  isMpvConnected: () => boolean;
  quitApp: () => void;
}) {
  return ({ connected }: { connected: boolean }): void => {
    deps.refreshDiscordPresence();
    if (connected) {
      deps.syncOverlayMpvSubtitleSuppression();
      return;
    }
    deps.reportJellyfinRemoteStopped();
    if (!deps.hasInitialPlaybackQuitOnDisconnectArg()) return;
    if (
      deps.isOverlayRuntimeInitialized() &&
      !deps.shouldQuitOnDisconnectWhenOverlayRuntimeInitialized()
    ) {
      return;
    }
    if (!deps.isQuitOnDisconnectArmed()) return;
    deps.scheduleQuitCheck(() => {
      if (deps.isMpvConnected()) return;
      deps.quitApp();
    });
  };
}

export function createHandleMpvSubtitleTimingHandler(deps: {
  recordImmersionSubtitleLine: (text: string, start: number, end: number) => void;
  hasSubtitleTimingTracker: () => boolean;
  recordSubtitleTiming: (text: string, start: number, end: number) => void;
  maybeRunAnilistPostWatchUpdate: (options?: AnilistPostWatchRunOptions) => Promise<void>;
  logError: (message: string, error: unknown) => void;
}) {
  return ({ text, start, end }: { text: string; start: number; end: number }): void => {
    const watchedSeconds = Math.max(
      Number.isFinite(start) ? start : 0,
      Number.isFinite(end) ? end : 0,
    );
    const options = watchedSeconds > 0 ? { watchedSeconds } : undefined;
    if (text.trim()) {
      deps.recordImmersionSubtitleLine(text, start, end);
      if (deps.hasSubtitleTimingTracker()) {
        deps.recordSubtitleTiming(text, start, end);
      }
    }
    void deps.maybeRunAnilistPostWatchUpdate(options).catch((error) => {
      deps.logError('AniList post-watch update failed unexpectedly', error);
    });
  };
}

export function createBindMpvClientEventHandlers(deps: {
  onConnectionChange: (payload: { connected: boolean }) => void;
  onSubtitleChange: (payload: { text: string }) => void;
  onSubtitleAssChange: (payload: { text: string }) => void;
  onSecondarySubtitleChange: (payload: { text: string }) => void;
  onSubtitleTrackChange: (payload: { sid: number | null }) => void;
  onSubtitleTrackListChange: (payload: { trackList: unknown[] | null }) => void;
  onSubtitleTiming: (payload: { text: string; start: number; end: number }) => void;
  onMediaPathChange: (payload: { path: string | null }) => void;
  onMediaTitleChange: (payload: { title: string | null }) => void;
  onTimePosChange: (payload: { time: number }) => void;
  onDurationChange: (payload: { duration: number }) => void;
  onPauseChange: (payload: { paused: boolean }) => void;
  onFullscreenChange: (payload: { fullscreen: boolean }) => void;
  onSubtitleMetricsChange: (payload: { patch: Record<string, unknown> }) => void;
  onSecondarySubtitleVisibility: (payload: { visible: boolean }) => void;
}) {
  return (mpvClient: MpvEventClient): void => {
    mpvClient.on('connection-change', deps.onConnectionChange);
    mpvClient.on('subtitle-change', deps.onSubtitleChange);
    mpvClient.on('subtitle-ass-change', deps.onSubtitleAssChange);
    mpvClient.on('secondary-subtitle-change', deps.onSecondarySubtitleChange);
    mpvClient.on('subtitle-track-change', deps.onSubtitleTrackChange);
    mpvClient.on('subtitle-track-list-change', deps.onSubtitleTrackListChange);
    mpvClient.on('subtitle-timing', deps.onSubtitleTiming);
    mpvClient.on('media-path-change', deps.onMediaPathChange);
    mpvClient.on('media-title-change', deps.onMediaTitleChange);
    mpvClient.on('time-pos-change', deps.onTimePosChange);
    mpvClient.on('duration-change', deps.onDurationChange);
    mpvClient.on('pause-change', deps.onPauseChange);
    mpvClient.on('fullscreen-change', deps.onFullscreenChange);
    mpvClient.on('subtitle-metrics-change', deps.onSubtitleMetricsChange);
    mpvClient.on('secondary-subtitle-visibility', deps.onSecondarySubtitleVisibility);
  };
}
