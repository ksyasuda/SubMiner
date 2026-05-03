import type { MergedToken, SubtitleData } from '../../types';

export function createBuildBindMpvMainEventHandlersMainDepsHandler(deps: {
  appState: {
    initialArgs?: {
      jellyfinPlay?: unknown;
      managedPlayback?: unknown;
      youtubePlay?: unknown;
    } | null;
    overlayRuntimeInitialized: boolean;
    mpvClient: {
      connected?: boolean;
      currentSecondarySubText?: string;
      currentTimePos?: number;
      requestProperty?: (name: string) => Promise<unknown>;
    } | null;
    immersionTracker: {
      recordSubtitleLine?: (
        text: string,
        start: number,
        end: number,
        tokens?: MergedToken[] | null,
        secondaryText?: string | null,
      ) => void;
      handleMediaTitleUpdate?: (title: string) => void;
      recordPlaybackPosition?: (time: number) => void;
      recordMediaDuration?: (durationSec: number) => void;
      recordPauseState?: (paused: boolean) => void;
    } | null;
    subtitleTimingTracker: {
      recordSubtitle?: (text: string, start: number, end: number) => void;
    } | null;
    currentMediaPath?: string | null;
    currentSubText: string;
    currentSubAssText: string;
    currentSubtitleData?: SubtitleData | null;
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
  getImmediateSubtitlePayload?: (text: string) => SubtitleData | null;
  emitImmediateSubtitle?: (payload: SubtitleData) => void;
  onSubtitleChange: (text: string) => void;
  onSubtitleTrackChange?: (sid: number | null) => void;
  onSubtitleTrackListChange?: (trackList: unknown[] | null) => void;
  updateCurrentMediaPath: (path: string) => void;
  restoreMpvSubVisibility: () => void;
  resetSubtitleSidebarEmbeddedLayout?: () => void;
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  maybeProbeAnilistDuration: (mediaKey: string) => void;
  ensureAnilistMediaGuess: (mediaKey: string) => void;
  syncImmersionMediaState: () => void;
  signalAutoplayReadyIfWarm?: (path: string) => void;
  scheduleCharacterDictionarySync?: () => void;
  updateCurrentMediaTitle: (title: string) => void;
  resetAnilistMediaGuessState: () => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  onTimePosUpdate?: (time: number) => void;
  onFullscreenChange?: (fullscreen: boolean) => void;
  updateSubtitleRenderMetrics: (patch: Record<string, unknown>) => void;
  refreshDiscordPresence: () => void;
  ensureImmersionTrackerInitialized: () => void;
  tokenizeSubtitleForImmersion?: (text: string) => Promise<SubtitleData | null>;
}) {
  const writePlaybackPositionFromMpv = (timeSec: unknown): void => {
    const normalizedTimeSec = Number(timeSec);
    if (!Number.isFinite(normalizedTimeSec)) {
      return;
    }
    deps.ensureImmersionTrackerInitialized();
    deps.appState.immersionTracker?.recordPlaybackPosition?.(normalizedTimeSec);
  };

  return () => ({
    reportJellyfinRemoteStopped: () => deps.reportJellyfinRemoteStopped(),
    syncOverlayMpvSubtitleSuppression: () => deps.syncOverlayMpvSubtitleSuppression(),
    hasInitialPlaybackQuitOnDisconnectArg: () =>
      Boolean(
        deps.appState.initialArgs?.managedPlayback ||
        deps.appState.initialArgs?.jellyfinPlay ||
        deps.appState.initialArgs?.youtubePlay,
      ),
    isOverlayRuntimeInitialized: () => deps.appState.overlayRuntimeInitialized,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () =>
      Boolean(
        deps.appState.initialArgs?.managedPlayback ||
        deps.appState.initialArgs?.jellyfinPlay ||
        deps.appState.initialArgs?.youtubePlay,
      ),
    isQuitOnDisconnectArmed: () => deps.getQuitOnDisconnectArmed(),
    scheduleQuitCheck: (callback: () => void) => deps.scheduleQuitCheck(callback),
    isMpvConnected: () => Boolean(deps.appState.mpvClient?.connected),
    quitApp: () => deps.quitApp(),
    recordImmersionSubtitleLine: (text: string, start: number, end: number) => {
      deps.ensureImmersionTrackerInitialized();
      const tracker = deps.appState.immersionTracker;
      if (!tracker?.recordSubtitleLine) {
        return;
      }
      const secondaryText = deps.appState.mpvClient?.currentSecondarySubText || null;
      const cachedTokens =
        deps.appState.currentSubtitleData?.text === text
          ? deps.appState.currentSubtitleData.tokens
          : null;
      if (cachedTokens) {
        tracker.recordSubtitleLine(text, start, end, cachedTokens, secondaryText);
        return;
      }
      if (!deps.tokenizeSubtitleForImmersion) {
        tracker.recordSubtitleLine(text, start, end, null, secondaryText);
        return;
      }
      void deps
        .tokenizeSubtitleForImmersion(text)
        .then((payload) => {
          tracker.recordSubtitleLine?.(text, start, end, payload?.tokens ?? null, secondaryText);
        })
        .catch(() => {
          tracker.recordSubtitleLine?.(text, start, end, null, secondaryText);
        });
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
    getImmediateSubtitlePayload: deps.getImmediateSubtitlePayload
      ? (text: string) => deps.getImmediateSubtitlePayload!(text)
      : undefined,
    emitImmediateSubtitle: deps.emitImmediateSubtitle
      ? (payload: SubtitleData) => deps.emitImmediateSubtitle!(payload)
      : undefined,
    broadcastSubtitle: (payload: SubtitleData) =>
      deps.broadcastToOverlayWindows('subtitle:set', payload),
    onSubtitleChange: (text: string) => deps.onSubtitleChange(text),
    onSubtitleTrackChange: deps.onSubtitleTrackChange
      ? (sid: number | null) => deps.onSubtitleTrackChange!(sid)
      : undefined,
    onSubtitleTrackListChange: deps.onSubtitleTrackListChange
      ? (trackList: unknown[] | null) => deps.onSubtitleTrackListChange!(trackList)
      : undefined,
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
    resetSubtitleSidebarEmbeddedLayout: () => deps.resetSubtitleSidebarEmbeddedLayout?.(),
    getCurrentAnilistMediaKey: () => deps.getCurrentAnilistMediaKey(),
    resetAnilistMediaTracking: (mediaKey: string | null) =>
      deps.resetAnilistMediaTracking(mediaKey),
    maybeProbeAnilistDuration: (mediaKey: string) => deps.maybeProbeAnilistDuration(mediaKey),
    ensureAnilistMediaGuess: (mediaKey: string) => deps.ensureAnilistMediaGuess(mediaKey),
    syncImmersionMediaState: () => deps.syncImmersionMediaState(),
    signalAutoplayReadyIfWarm: (path: string) => deps.signalAutoplayReadyIfWarm?.(path),
    scheduleCharacterDictionarySync: () => deps.scheduleCharacterDictionarySync?.(),
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
    recordMediaDuration: (durationSec: number) => {
      deps.ensureImmersionTrackerInitialized();
      deps.appState.immersionTracker?.recordMediaDuration?.(durationSec);
    },
    reportJellyfinRemoteProgress: (forceImmediate: boolean) =>
      deps.reportJellyfinRemoteProgress(forceImmediate),
    onTimePosUpdate: deps.onTimePosUpdate
      ? (time: number) => deps.onTimePosUpdate!(time)
      : undefined,
    onFullscreenChange: deps.onFullscreenChange
      ? (fullscreen: boolean) => deps.onFullscreenChange!(fullscreen)
      : undefined,
    recordPauseState: (paused: boolean) => {
      deps.appState.playbackPaused = paused;
      deps.ensureImmersionTrackerInitialized();
      deps.appState.immersionTracker?.recordPauseState?.(paused);
    },
    flushPlaybackPositionOnMediaPathClear: (mediaPath: string) => {
      const mpvClient = deps.appState.mpvClient;
      const currentKnownTime = Number(mpvClient?.currentTimePos);
      writePlaybackPositionFromMpv(currentKnownTime);
      if (!mpvClient?.requestProperty) {
        return;
      }
      void mpvClient
        .requestProperty('time-pos')
        .then((timePos) => {
          const currentPath = (deps.appState.currentMediaPath ?? '').trim();
          if (currentPath.length > 0 && currentPath !== mediaPath) {
            return;
          }
          const resolvedTime = Number(timePos);
          if (
            Number.isFinite(currentKnownTime) &&
            Number.isFinite(resolvedTime) &&
            currentKnownTime === resolvedTime
          ) {
            return;
          }
          writePlaybackPositionFromMpv(resolvedTime);
        })
        .catch(() => {
          // mpv can disconnect while clearing media; keep the last cached position.
        });
    },
    updateSubtitleRenderMetrics: (patch: Record<string, unknown>) =>
      deps.updateSubtitleRenderMetrics(patch),
    setPreviousSecondarySubVisibility: (visible: boolean) => {
      deps.appState.previousSecondarySubVisibility = visible;
    },
  });
}
