export function createBuildBindMpvMainEventHandlersMainDepsHandler(deps: {
  appState: {
    initialArgs?: { jellyfinPlay?: unknown } | null;
    overlayRuntimeInitialized: boolean;
    mpvClient: { connected?: boolean } | null;
    immersionTracker: {
      recordSubtitleLine?: (text: string, start: number, end: number) => void;
      handleMediaTitleUpdate?: (title: string) => void;
      recordPlaybackPosition?: (time: number) => void;
      recordPauseState?: (paused: boolean) => void;
    } | null;
    subtitleTimingTracker: {
      recordSubtitle?: (text: string, start: number, end: number) => void;
    } | null;
    currentSubText: string;
    currentSubAssText: string;
    playbackPaused: boolean | null;
    previousSecondarySubVisibility: boolean | null;
  };
  getQuitOnDisconnectArmed: () => boolean;
  scheduleQuitCheck: (callback: () => void) => void;
  quitApp: () => void;
  reportJellyfinRemoteStopped: () => void;
  syncOverlayMpvSubtitleSuppression: () => void;
  maybeRunAnilistPostWatchUpdate: () => Promise<void>;
  logSubtitleTimingError: (message: string, error: unknown) => void;
  broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  onSubtitleChange: (text: string) => void;
  updateCurrentMediaPath: (path: string) => void;
  restoreMpvSubVisibility: () => void;
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  maybeProbeAnilistDuration: (mediaKey: string) => void;
  ensureAnilistMediaGuess: (mediaKey: string) => void;
  syncImmersionMediaState: () => void;
  updateCurrentMediaTitle: (title: string) => void;
  resetAnilistMediaGuessState: () => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  updateSubtitleRenderMetrics: (patch: Record<string, unknown>) => void;
  refreshDiscordPresence: () => void;
  ensureImmersionTrackerInitialized: () => void;
}) {
  return () => ({
    reportJellyfinRemoteStopped: () => deps.reportJellyfinRemoteStopped(),
    syncOverlayMpvSubtitleSuppression: () => deps.syncOverlayMpvSubtitleSuppression(),
    hasInitialJellyfinPlayArg: () => Boolean(deps.appState.initialArgs?.jellyfinPlay),
    isOverlayRuntimeInitialized: () => deps.appState.overlayRuntimeInitialized,
    isQuitOnDisconnectArmed: () => deps.getQuitOnDisconnectArmed(),
    scheduleQuitCheck: (callback: () => void) => deps.scheduleQuitCheck(callback),
    isMpvConnected: () => Boolean(deps.appState.mpvClient?.connected),
    quitApp: () => deps.quitApp(),
    recordImmersionSubtitleLine: (text: string, start: number, end: number) => {
      deps.ensureImmersionTrackerInitialized();
      deps.appState.immersionTracker?.recordSubtitleLine?.(text, start, end);
    },
    hasSubtitleTimingTracker: () => Boolean(deps.appState.subtitleTimingTracker),
    recordSubtitleTiming: (text: string, start: number, end: number) =>
      deps.appState.subtitleTimingTracker?.recordSubtitle?.(text, start, end),
    maybeRunAnilistPostWatchUpdate: () => deps.maybeRunAnilistPostWatchUpdate(),
    logSubtitleTimingError: (message: string, error: unknown) =>
      deps.logSubtitleTimingError(message, error),
    setCurrentSubText: (text: string) => {
      deps.appState.currentSubText = text;
    },
    broadcastSubtitle: (payload: { text: string; tokens: null }) =>
      deps.broadcastToOverlayWindows('subtitle:set', payload),
    onSubtitleChange: (text: string) => deps.onSubtitleChange(text),
    refreshDiscordPresence: () => deps.refreshDiscordPresence(),
    setCurrentSubAssText: (text: string) => {
      deps.appState.currentSubAssText = text;
    },
    broadcastSubtitleAss: (text: string) =>
      deps.broadcastToOverlayWindows('subtitle-ass:set', text),
    broadcastSecondarySubtitle: (text: string) =>
      deps.broadcastToOverlayWindows('secondary-subtitle:set', text),
    updateCurrentMediaPath: (path: string) => deps.updateCurrentMediaPath(path),
    restoreMpvSubVisibility: () => deps.restoreMpvSubVisibility(),
    getCurrentAnilistMediaKey: () => deps.getCurrentAnilistMediaKey(),
    resetAnilistMediaTracking: (mediaKey: string | null) =>
      deps.resetAnilistMediaTracking(mediaKey),
    maybeProbeAnilistDuration: (mediaKey: string) => deps.maybeProbeAnilistDuration(mediaKey),
    ensureAnilistMediaGuess: (mediaKey: string) => deps.ensureAnilistMediaGuess(mediaKey),
    syncImmersionMediaState: () => deps.syncImmersionMediaState(),
    updateCurrentMediaTitle: (title: string) => deps.updateCurrentMediaTitle(title),
    resetAnilistMediaGuessState: () => deps.resetAnilistMediaGuessState(),
    notifyImmersionTitleUpdate: (title: string) => {
      deps.ensureImmersionTrackerInitialized();
      deps.appState.immersionTracker?.handleMediaTitleUpdate?.(title);
    },
    recordPlaybackPosition: (time: number) => {
      deps.ensureImmersionTrackerInitialized();
      deps.appState.immersionTracker?.recordPlaybackPosition?.(time);
    },
    reportJellyfinRemoteProgress: (forceImmediate: boolean) =>
      deps.reportJellyfinRemoteProgress(forceImmediate),
    recordPauseState: (paused: boolean) => {
      deps.appState.playbackPaused = paused;
      deps.ensureImmersionTrackerInitialized();
      deps.appState.immersionTracker?.recordPauseState?.(paused);
    },
    updateSubtitleRenderMetrics: (patch: Record<string, unknown>) =>
      deps.updateSubtitleRenderMetrics(patch),
    setPreviousSecondarySubVisibility: (visible: boolean) => {
      deps.appState.previousSecondarySubVisibility = visible;
    },
  });
}
