import {
  createBindMpvClientEventHandlers,
  createHandleMpvConnectionChangeHandler,
  createHandleMpvSubtitleTimingHandler,
} from './mpv-client-event-bindings';
import {
  createHandleMpvMediaPathChangeHandler,
  createHandleMpvMediaTitleChangeHandler,
  createHandleMpvPauseChangeHandler,
  createHandleMpvSecondarySubtitleChangeHandler,
  createHandleMpvSecondarySubtitleVisibilityHandler,
  createHandleMpvSubtitleAssChangeHandler,
  createHandleMpvSubtitleChangeHandler,
  createHandleMpvSubtitleMetricsChangeHandler,
  createHandleMpvTimePosChangeHandler,
} from './mpv-main-event-actions';

type MpvEventClient = Parameters<ReturnType<typeof createBindMpvClientEventHandlers>>[0];

export function createBindMpvMainEventHandlersHandler(deps: {
  reportJellyfinRemoteStopped: () => void;
  hasInitialJellyfinPlayArg: () => boolean;
  isOverlayRuntimeInitialized: () => boolean;
  isQuitOnDisconnectArmed: () => boolean;
  scheduleQuitCheck: (callback: () => void) => void;
  isMpvConnected: () => boolean;
  quitApp: () => void;

  recordImmersionSubtitleLine: (text: string, start: number, end: number) => void;
  hasSubtitleTimingTracker: () => boolean;
  recordSubtitleTiming: (text: string, start: number, end: number) => void;
  maybeRunAnilistPostWatchUpdate: () => Promise<void>;
  logSubtitleTimingError: (message: string, error: unknown) => void;

  setCurrentSubText: (text: string) => void;
  broadcastSubtitle: (payload: { text: string; tokens: null }) => void;
  onSubtitleChange: (text: string) => void;

  setCurrentSubAssText: (text: string) => void;
  broadcastSubtitleAss: (text: string) => void;
  broadcastSecondarySubtitle: (text: string) => void;

  updateCurrentMediaPath: (path: string) => void;
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  maybeProbeAnilistDuration: (mediaKey: string) => void;
  ensureAnilistMediaGuess: (mediaKey: string) => void;
  syncImmersionMediaState: () => void;

  updateCurrentMediaTitle: (title: string) => void;
  resetAnilistMediaGuessState: () => void;
  notifyImmersionTitleUpdate: (title: string) => void;

  recordPlaybackPosition: (time: number) => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  recordPauseState: (paused: boolean) => void;

  updateSubtitleRenderMetrics: (patch: Record<string, unknown>) => void;
  setPreviousSecondarySubVisibility: (visible: boolean) => void;
}) {
  return (mpvClient: MpvEventClient): void => {
    const handleMpvConnectionChange = createHandleMpvConnectionChangeHandler({
      reportJellyfinRemoteStopped: () => deps.reportJellyfinRemoteStopped(),
      hasInitialJellyfinPlayArg: () => deps.hasInitialJellyfinPlayArg(),
      isOverlayRuntimeInitialized: () => deps.isOverlayRuntimeInitialized(),
      isQuitOnDisconnectArmed: () => deps.isQuitOnDisconnectArmed(),
      scheduleQuitCheck: (callback) => deps.scheduleQuitCheck(callback),
      isMpvConnected: () => deps.isMpvConnected(),
      quitApp: () => deps.quitApp(),
    });
    const handleMpvSubtitleTiming = createHandleMpvSubtitleTimingHandler({
      recordImmersionSubtitleLine: (text, start, end) =>
        deps.recordImmersionSubtitleLine(text, start, end),
      hasSubtitleTimingTracker: () => deps.hasSubtitleTimingTracker(),
      recordSubtitleTiming: (text, start, end) => deps.recordSubtitleTiming(text, start, end),
      maybeRunAnilistPostWatchUpdate: () => deps.maybeRunAnilistPostWatchUpdate(),
      logError: (message, error) => deps.logSubtitleTimingError(message, error),
    });
    const handleMpvSubtitleChange = createHandleMpvSubtitleChangeHandler({
      setCurrentSubText: (text) => deps.setCurrentSubText(text),
      broadcastSubtitle: (payload) => deps.broadcastSubtitle(payload),
      onSubtitleChange: (text) => deps.onSubtitleChange(text),
    });
    const handleMpvSubtitleAssChange = createHandleMpvSubtitleAssChangeHandler({
      setCurrentSubAssText: (text) => deps.setCurrentSubAssText(text),
      broadcastSubtitleAss: (text) => deps.broadcastSubtitleAss(text),
    });
    const handleMpvSecondarySubtitleChange = createHandleMpvSecondarySubtitleChangeHandler({
      broadcastSecondarySubtitle: (text) => deps.broadcastSecondarySubtitle(text),
    });
    const handleMpvMediaPathChange = createHandleMpvMediaPathChangeHandler({
      updateCurrentMediaPath: (path) => deps.updateCurrentMediaPath(path),
      reportJellyfinRemoteStopped: () => deps.reportJellyfinRemoteStopped(),
      getCurrentAnilistMediaKey: () => deps.getCurrentAnilistMediaKey(),
      resetAnilistMediaTracking: (mediaKey) => deps.resetAnilistMediaTracking(mediaKey),
      maybeProbeAnilistDuration: (mediaKey) => deps.maybeProbeAnilistDuration(mediaKey),
      ensureAnilistMediaGuess: (mediaKey) => deps.ensureAnilistMediaGuess(mediaKey),
      syncImmersionMediaState: () => deps.syncImmersionMediaState(),
    });
    const handleMpvMediaTitleChange = createHandleMpvMediaTitleChangeHandler({
      updateCurrentMediaTitle: (title) => deps.updateCurrentMediaTitle(title),
      resetAnilistMediaGuessState: () => deps.resetAnilistMediaGuessState(),
      notifyImmersionTitleUpdate: (title) => deps.notifyImmersionTitleUpdate(title),
      syncImmersionMediaState: () => deps.syncImmersionMediaState(),
    });
    const handleMpvTimePosChange = createHandleMpvTimePosChangeHandler({
      recordPlaybackPosition: (time) => deps.recordPlaybackPosition(time),
      reportJellyfinRemoteProgress: (forceImmediate) =>
        deps.reportJellyfinRemoteProgress(forceImmediate),
    });
    const handleMpvPauseChange = createHandleMpvPauseChangeHandler({
      recordPauseState: (paused) => deps.recordPauseState(paused),
      reportJellyfinRemoteProgress: (forceImmediate) =>
        deps.reportJellyfinRemoteProgress(forceImmediate),
    });
    const handleMpvSubtitleMetricsChange = createHandleMpvSubtitleMetricsChangeHandler({
      updateSubtitleRenderMetrics: (patch) => deps.updateSubtitleRenderMetrics(patch),
    });
    const handleMpvSecondarySubtitleVisibility = createHandleMpvSecondarySubtitleVisibilityHandler({
      setPreviousSecondarySubVisibility: (visible) =>
        deps.setPreviousSecondarySubVisibility(visible),
    });

    createBindMpvClientEventHandlers({
      onConnectionChange: handleMpvConnectionChange,
      onSubtitleChange: handleMpvSubtitleChange,
      onSubtitleAssChange: handleMpvSubtitleAssChange,
      onSecondarySubtitleChange: handleMpvSecondarySubtitleChange,
      onSubtitleTiming: handleMpvSubtitleTiming,
      onMediaPathChange: handleMpvMediaPathChange,
      onMediaTitleChange: handleMpvMediaTitleChange,
      onTimePosChange: handleMpvTimePosChange,
      onPauseChange: handleMpvPauseChange,
      onSubtitleMetricsChange: handleMpvSubtitleMetricsChange,
      onSecondarySubtitleVisibility: handleMpvSecondarySubtitleVisibility,
    })(mpvClient);
  };
}
