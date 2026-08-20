import type { SubtitleData } from '../../types';
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

type AnilistPostWatchRunOptions = {
  watchedSeconds?: number;
};

export function createBindMpvMainEventHandlersHandler(deps: {
  reportJellyfinRemoteStopped: () => void;
  syncOverlayMpvSubtitleSuppression: () => void;
  onMpvConnected?: () => void;
  resetSubtitleSidebarEmbeddedLayout: () => void;
  scheduleCharacterDictionarySync?: () => void;
  hasInitialPlaybackQuitOnDisconnectArg: () => boolean;
  isOverlayRuntimeInitialized: () => boolean;
  shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => boolean;
  isQuitOnDisconnectArmed: () => boolean;
  scheduleQuitCheck: (callback: () => void) => void;
  isMpvConnected: () => boolean;
  quitApp: () => void;

  recordImmersionSubtitleLine: (text: string, start: number, end: number) => void;
  hasSubtitleTimingTracker: () => boolean;
  recordSubtitleTiming: (text: string, start: number, end: number) => void;
  maybeRunAnilistPostWatchUpdate: (options?: AnilistPostWatchRunOptions) => Promise<void>;
  logSubtitleTimingError: (message: string, error: unknown) => void;

  setCurrentSubText: (text: string) => void;
  resolveSubtitleText?: (text: string) => string;
  getCurrentLiveSubtitleText?: () => string;
  getImmediateSubtitlePayload?: (text: string) => SubtitleData | null;
  emitImmediateSubtitle?: (payload: SubtitleData) => void;
  broadcastSubtitle: (payload: SubtitleData) => void;
  onSubtitleChange: (text: string) => void;
  logSubtitleProcessingDebug?: (message: string) => void;
  refreshDiscordPresence: () => void;

  setCurrentSubAssText: (text: string) => void;
  broadcastSubtitleAss: (text: string) => void;
  broadcastSecondarySubtitle: (text: string) => void;
  onSubtitleTrackChange?: (sid: number | null) => void;
  onSecondarySubtitleTrackChange?: (sid: number | null) => void;
  onSecondarySubtitleDelayChange?: (delay: number) => void;
  onSubtitleTrackListChange?: (trackList: unknown[] | null) => void;

  updateCurrentMediaPath: (path: string) => void;
  restoreMpvSubVisibility: () => void;
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  maybeProbeAnilistDuration: (mediaKey: string) => void;
  ensureAnilistMediaGuess: (mediaKey: string) => void;
  syncImmersionMediaState: () => void;
  signalAutoplayReadyIfWarm?: (path: string) => void;
  markJellyfinRemotePlaybackLoaded?: (path: string) => void;
  flushPlaybackPositionOnMediaPathClear?: (mediaPath: string) => void;

  updateCurrentMediaTitle: (title: string) => void;
  resetAnilistMediaGuessState: () => void;
  notifyImmersionTitleUpdate: (title: string) => void;

  recordPlaybackPosition: (time: number) => void;
  recordMediaDuration: (durationSec: number) => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  onTimePosUpdate?: (time: number) => void;
  consumeExplicitSeek?: () => boolean;
  onFullscreenChange?: (fullscreen: boolean) => void;
  recordPauseState: (paused: boolean) => void;

  updateSubtitleRenderMetrics: (patch: Record<string, unknown>) => void;
  setPreviousSecondarySubVisibility: (visible: boolean) => void;
}) {
  return (mpvClient: MpvEventClient): void => {
    const handleMpvConnectionChange = createHandleMpvConnectionChangeHandler({
      reportJellyfinRemoteStopped: () => deps.reportJellyfinRemoteStopped(),
      refreshDiscordPresence: () => deps.refreshDiscordPresence(),
      syncOverlayMpvSubtitleSuppression: () => deps.syncOverlayMpvSubtitleSuppression(),
      onConnected: () => deps.onMpvConnected?.(),
      hasInitialPlaybackQuitOnDisconnectArg: () => deps.hasInitialPlaybackQuitOnDisconnectArg(),
      isOverlayRuntimeInitialized: () => deps.isOverlayRuntimeInitialized(),
      shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () =>
        deps.shouldQuitOnDisconnectWhenOverlayRuntimeInitialized(),
      isQuitOnDisconnectArmed: () => deps.isQuitOnDisconnectArmed(),
      scheduleQuitCheck: (callback) => deps.scheduleQuitCheck(callback),
      isMpvConnected: () => deps.isMpvConnected(),
      quitApp: () => deps.quitApp(),
    });
    const handleMpvConnectionChangeWithSidebarReset = ({
      connected,
    }: {
      connected: boolean;
    }): void => {
      if (connected) {
        deps.resetSubtitleSidebarEmbeddedLayout();
      } else {
        deps.updateCurrentMediaPath('');
      }
      handleMpvConnectionChange({ connected });
    };
    const handleMpvSubtitleTiming = createHandleMpvSubtitleTimingHandler({
      recordImmersionSubtitleLine: (text, start, end) =>
        deps.recordImmersionSubtitleLine(text, start, end),
      hasSubtitleTimingTracker: () => deps.hasSubtitleTimingTracker(),
      recordSubtitleTiming: (text, start, end) => deps.recordSubtitleTiming(text, start, end),
      maybeRunAnilistPostWatchUpdate: (options) => deps.maybeRunAnilistPostWatchUpdate(options),
      logError: (message, error) => deps.logSubtitleTimingError(message, error),
    });
    const handleMpvSubtitleChange = createHandleMpvSubtitleChangeHandler({
      resolveSubtitleText: deps.resolveSubtitleText,
      setCurrentSubText: (text) => deps.setCurrentSubText(text),
      getImmediateSubtitlePayload: (text) => deps.getImmediateSubtitlePayload?.(text) ?? null,
      emitImmediateSubtitle: deps.emitImmediateSubtitle
        ? (payload) => deps.emitImmediateSubtitle?.(payload)
        : undefined,
      broadcastSubtitle: (payload) => deps.broadcastSubtitle(payload),
      onSubtitleChange: (text) => deps.onSubtitleChange(text),
      logDebug: deps.logSubtitleProcessingDebug
        ? (message) => deps.logSubtitleProcessingDebug?.(message)
        : undefined,
      refreshDiscordPresence: () => deps.refreshDiscordPresence(),
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
      restoreMpvSubVisibility: () => deps.restoreMpvSubVisibility(),
      resetSubtitleSidebarEmbeddedLayout: () => deps.resetSubtitleSidebarEmbeddedLayout(),
      getCurrentAnilistMediaKey: () => deps.getCurrentAnilistMediaKey(),
      resetAnilistMediaTracking: (mediaKey) => deps.resetAnilistMediaTracking(mediaKey),
      maybeProbeAnilistDuration: (mediaKey) => deps.maybeProbeAnilistDuration(mediaKey),
      ensureAnilistMediaGuess: (mediaKey) => deps.ensureAnilistMediaGuess(mediaKey),
      syncImmersionMediaState: () => deps.syncImmersionMediaState(),
      flushPlaybackPositionOnMediaPathClear: (mediaPath) =>
        deps.flushPlaybackPositionOnMediaPathClear?.(mediaPath),
      signalAutoplayReadyIfWarm: (path) => deps.signalAutoplayReadyIfWarm?.(path),
      markJellyfinRemotePlaybackLoaded: (path) => deps.markJellyfinRemotePlaybackLoaded?.(path),
      scheduleCharacterDictionarySync: () => deps.scheduleCharacterDictionarySync?.(),
      refreshDiscordPresence: () => deps.refreshDiscordPresence(),
    });
    const handleMpvMediaTitleChange = createHandleMpvMediaTitleChangeHandler({
      updateCurrentMediaTitle: (title) => deps.updateCurrentMediaTitle(title),
      resetAnilistMediaGuessState: () => deps.resetAnilistMediaGuessState(),
      notifyImmersionTitleUpdate: (title) => deps.notifyImmersionTitleUpdate(title),
      syncImmersionMediaState: () => deps.syncImmersionMediaState(),
      refreshDiscordPresence: () => deps.refreshDiscordPresence(),
    });
    const handleMpvTimePosChange = createHandleMpvTimePosChangeHandler({
      recordPlaybackPosition: (time) => deps.recordPlaybackPosition(time),
      reportJellyfinRemoteProgress: (forceImmediate) =>
        deps.reportJellyfinRemoteProgress(forceImmediate),
      refreshDiscordPresence: () => deps.refreshDiscordPresence(),
      maybeRunAnilistPostWatchUpdate: (options) => deps.maybeRunAnilistPostWatchUpdate(options),
      logError: (message, error) => deps.logSubtitleTimingError(message, error),
      consumeExplicitSeek: deps.consumeExplicitSeek,
      onTimePosUpdate: (time, updateKind) => {
        deps.onTimePosUpdate?.(time);
        if (updateKind === 'playback') return;
        const liveText = deps.getCurrentLiveSubtitleText?.();
        if (liveText !== undefined) {
          handleMpvSubtitleChange({ text: liveText });
        }
      },
    });
    const handleMpvPauseChange = createHandleMpvPauseChangeHandler({
      recordPauseState: (paused) => deps.recordPauseState(paused),
      reportJellyfinRemoteProgress: (forceImmediate) =>
        deps.reportJellyfinRemoteProgress(forceImmediate),
      refreshDiscordPresence: () => deps.refreshDiscordPresence(),
    });
    const handleMpvSubtitleMetricsChange = createHandleMpvSubtitleMetricsChangeHandler({
      updateSubtitleRenderMetrics: (patch) => deps.updateSubtitleRenderMetrics(patch),
    });
    const handleMpvSecondarySubtitleVisibility = createHandleMpvSecondarySubtitleVisibilityHandler({
      setPreviousSecondarySubVisibility: (visible) =>
        deps.setPreviousSecondarySubVisibility(visible),
    });

    createBindMpvClientEventHandlers({
      onConnectionChange: handleMpvConnectionChangeWithSidebarReset,
      onSubtitleChange: handleMpvSubtitleChange,
      onSubtitleAssChange: handleMpvSubtitleAssChange,
      onSecondarySubtitleChange: handleMpvSecondarySubtitleChange,
      onSubtitleTrackChange: ({ sid }) => deps.onSubtitleTrackChange?.(sid),
      onSecondarySubtitleTrackChange: ({ sid }) => deps.onSecondarySubtitleTrackChange?.(sid),
      onSecondarySubtitleDelayChange: ({ delay }) => deps.onSecondarySubtitleDelayChange?.(delay),
      onSubtitleTrackListChange: ({ trackList }) => deps.onSubtitleTrackListChange?.(trackList),
      onSubtitleTiming: handleMpvSubtitleTiming,
      onMediaPathChange: handleMpvMediaPathChange,
      onMediaTitleChange: handleMpvMediaTitleChange,
      onTimePosChange: handleMpvTimePosChange,
      onDurationChange: ({ duration }) => deps.recordMediaDuration(duration),
      onPauseChange: handleMpvPauseChange,
      onFullscreenChange: ({ fullscreen }) => deps.onFullscreenChange?.(fullscreen),
      onSubtitleMetricsChange: handleMpvSubtitleMetricsChange,
      onSecondarySubtitleVisibility: handleMpvSecondarySubtitleVisibility,
    })(mpvClient);
  };
}
