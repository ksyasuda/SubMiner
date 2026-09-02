import { createSubtitleLineDedupGate } from '../../core/services/subtitle-line-dedup-gate';
import type { MergedToken, SubtitleCue, SubtitleData } from '../../types';
import { SEEK_LIKE_TIME_DELTA_SECONDS } from './mpv-main-event-actions';
import {
  resolveCanonicalPrimarySubtitle,
  resolveParsedPrimarySubtitle,
  resolvePrimarySubtitleText,
  stripCanonicalFragmentLines,
} from './primary-subtitle-text';

type AnilistPostWatchRunOptions = {
  watchedSeconds?: number;
};

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
      currentSubText?: string;
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
      recordSubtitle?: (text: string, start: number, end: number, secondaryText?: string) => void;
    } | null;
    activeParsedSubtitleCues?: SubtitleCue[] | null;
    /** Cache key of the source the cues were parsed from; cleared with the cues. */
    activeParsedSubtitleSource?: string | null;
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
  onMpvConnected?: () => void;
  maybeRunAnilistPostWatchUpdate: (options?: AnilistPostWatchRunOptions) => Promise<void>;
  recordAnilistMediaDuration?: (durationSec: number) => void;
  logSubtitleTimingError: (message: string, error: unknown) => void;
  broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  onSecondarySubtitleChange?: (text: string) => void;
  getImmediateSubtitlePayload?: (text: string) => SubtitleData | null;
  emitImmediateSubtitle?: (payload: SubtitleData) => void;
  onSubtitleChange: (text: string) => void;
  logSubtitleProcessingDebug?: (message: string) => void;
  onSubtitleTrackChange?: (sid: number | null) => void;
  onSecondarySubtitleTrackChange?: (sid: number | null) => void;
  onSecondarySubtitleDelayChange?: (delay: number) => void;
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
  markJellyfinRemotePlaybackLoaded?: (path: string) => void;
  scheduleCharacterDictionarySync?: () => void;
  updateCurrentMediaTitle: (title: string) => void;
  resetAnilistMediaGuessState: () => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  onTimePosUpdate?: (time: number) => void;
  consumeExplicitSeek?: () => boolean;
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
  // mpv reports every animation frame of a typeset line as its own subtitle event, so
  // stats have to collapse bursts the same way the parsed cue list already does.
  const immersionLineDedupGate = createSubtitleLineDedupGate({
    getParsedCues: () => deps.appState.activeParsedSubtitleCues,
  });
  // One seen-set per consumer: canonical cues overlap, so live samples resolve to
  // shifting subsets (A, then A+B, then B). Remembering every recorded cue -- not just
  // the previous sample -- keeps each authored line recorded exactly once per source.
  const recordedImmersionCanonicalKeys = new Set<string>();
  const recordedTimingCanonicalKeys = new Set<string>();
  // Bumped on track/media changes so an immersion record whose tokenization resolves
  // after the change is dropped instead of landing in the next session.
  let subtitleSessionEpoch = 0;
  let lastTimePosForTimingReset: number | null = null;
  const canonicalCueKey = (cue: SubtitleCue): string =>
    `${cue.startTime}|${cue.endTime}|${cue.text}`;
  const resetSubtitleDeduplication = (): void => {
    immersionLineDedupGate.reset();
    recordedImmersionCanonicalKeys.clear();
    recordedTimingCanonicalKeys.clear();
    subtitleSessionEpoch += 1;
    lastTimePosForTimingReset = null;
  };
  const resolveCanonicalSample = (liveText: string, startSec: number) =>
    resolveCanonicalPrimarySubtitle({
      liveText,
      currentTimeSec: startSec,
      cues: deps.appState.activeParsedSubtitleCues,
    });
  // Recorders see the same text the overlay displays: mpv's live `sub-text` lists every
  // simultaneously active ASS event, including furigana events the parser folded into
  // their base line, so substitute the parsed view wherever it explains every live line.
  // Canonical substitution is the callers' own branch, so what is left here is the parsed
  // view, and its text is already a complete line -- fragment stripping takes raw mpv
  // text and would discard a resolved line whole while a fragment grid is on screen. When
  // nothing explained the sample -- dialogue sharing the screen with a song -- strip the
  // raw text so the dialogue is recorded without the fragments beside it.
  const resolveTextForRecording = (liveText: string, startSec: number): string => {
    const cues = deps.appState.activeParsedSubtitleCues;
    return (
      resolveParsedPrimarySubtitle({ liveText, currentTimeSec: startSec, cues })?.text ??
      stripCanonicalFragmentLines({ liveText, currentTimeSec: startSec, cues })
    );
  };
  const hasInitialPlaybackQuitOnDisconnectArg = (): boolean =>
    Boolean(
      deps.appState.initialArgs?.managedPlayback ||
      deps.appState.initialArgs?.jellyfinPlay ||
      deps.appState.initialArgs?.youtubePlay,
    );

  return () => ({
    reportJellyfinRemoteStopped: () => deps.reportJellyfinRemoteStopped(),
    syncOverlayMpvSubtitleSuppression: () => deps.syncOverlayMpvSubtitleSuppression(),
    onMpvConnected: deps.onMpvConnected ? () => deps.onMpvConnected!() : undefined,
    hasInitialPlaybackQuitOnDisconnectArg,
    isOverlayRuntimeInitialized: () => deps.appState.overlayRuntimeInitialized,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: hasInitialPlaybackQuitOnDisconnectArg,
    isQuitOnDisconnectArmed: () => deps.getQuitOnDisconnectArmed(),
    scheduleQuitCheck: (callback: () => void) => deps.scheduleQuitCheck(callback),
    isMpvConnected: () => Boolean(deps.appState.mpvClient?.connected),
    quitApp: () => deps.quitApp(),
    resolveSubtitleText: (liveText: string) =>
      resolvePrimarySubtitleText({
        liveText,
        currentTimeSec: Number(deps.appState.mpvClient?.currentTimePos),
        cues: deps.appState.activeParsedSubtitleCues,
      }),
    getCurrentLiveSubtitleText: () => deps.appState.mpvClient?.currentSubText ?? '',
    recordImmersionSubtitleLine: (text: string, start: number, end: number) => {
      deps.ensureImmersionTrackerInitialized();
      const tracker = deps.appState.immersionTracker;
      if (!tracker?.recordSubtitleLine) {
        return;
      }
      const recordLine = (lineText: string, startSec: number, endSec: number): void => {
        const secondaryText = deps.appState.mpvClient?.currentSecondarySubText || null;
        const cachedTokens =
          deps.appState.currentSubtitleData?.text === lineText
            ? deps.appState.currentSubtitleData.tokens
            : null;
        if (cachedTokens) {
          tracker.recordSubtitleLine?.(lineText, startSec, endSec, cachedTokens, secondaryText);
          return;
        }
        if (!deps.tokenizeSubtitleForImmersion) {
          tracker.recordSubtitleLine?.(lineText, startSec, endSec, null, secondaryText);
          return;
        }
        const epochAtRecord = subtitleSessionEpoch;
        void deps
          .tokenizeSubtitleForImmersion(lineText)
          .then((payload) => {
            if (subtitleSessionEpoch !== epochAtRecord) {
              return;
            }
            tracker.recordSubtitleLine?.(
              lineText,
              startSec,
              endSec,
              payload?.tokens ?? null,
              secondaryText,
            );
          })
          .catch(() => {
            if (subtitleSessionEpoch !== epochAtRecord) {
              return;
            }
            tracker.recordSubtitleLine?.(lineText, startSec, endSec, null, secondaryText);
          });
      };
      const canonical = resolveCanonicalSample(text, start);
      if (canonical) {
        for (const cue of canonical.cues) {
          const key = canonicalCueKey(cue);
          if (recordedImmersionCanonicalKeys.has(key)) {
            continue;
          }
          recordedImmersionCanonicalKeys.add(key);
          recordLine(cue.text, cue.startTime, cue.endTime);
        }
        return;
      }
      text = resolveTextForRecording(text, start);
      if (!text.trim()) {
        return;
      }
      if (!immersionLineDedupGate.shouldRecord({ text, startSec: start, endSec: end })) {
        return;
      }
      recordLine(text, start, end);
    },
    hasSubtitleTimingTracker: () => Boolean(deps.appState.subtitleTimingTracker),
    recordSubtitleTiming: (text: string, start: number, end: number) => {
      const secondaryText = deps.appState.mpvClient?.currentSecondarySubText || undefined;
      const canonical = resolveCanonicalSample(text, start);
      if (!canonical) {
        const recordableText = resolveTextForRecording(text, start);
        if (!recordableText.trim()) {
          return;
        }
        deps.appState.subtitleTimingTracker?.recordSubtitle?.(
          recordableText,
          start,
          end,
          secondaryText,
        );
        return;
      }
      for (const cue of canonical.cues) {
        const key = canonicalCueKey(cue);
        if (recordedTimingCanonicalKeys.has(key)) {
          continue;
        }
        recordedTimingCanonicalKeys.add(key);
        deps.appState.subtitleTimingTracker?.recordSubtitle?.(
          cue.text,
          cue.startTime,
          cue.endTime,
          secondaryText,
        );
      }
    },
    maybeRunAnilistPostWatchUpdate: (options?: AnilistPostWatchRunOptions) =>
      deps.maybeRunAnilistPostWatchUpdate(options),
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
    logSubtitleProcessingDebug: deps.logSubtitleProcessingDebug
      ? (message: string) => deps.logSubtitleProcessingDebug!(message)
      : undefined,
    onSubtitleTrackChange: (sid: number | null) => {
      resetSubtitleDeduplication();
      // The replacement track's cues arrive only after an async re-read and re-parse.
      // Clearing synchronously keeps the previous track's canonical cues from
      // substituting into, or recording against, the new track's live text. The source
      // key is cleared with the cues so cue-list consumers (the sidebar snapshot)
      // re-parse on demand instead of trusting the stale pairing.
      deps.appState.activeParsedSubtitleCues = [];
      deps.appState.activeParsedSubtitleSource = null;
      deps.onSubtitleTrackChange?.(sid);
    },
    onSecondarySubtitleTrackChange: deps.onSecondarySubtitleTrackChange
      ? (sid: number | null) => deps.onSecondarySubtitleTrackChange!(sid)
      : undefined,
    onSecondarySubtitleDelayChange: deps.onSecondarySubtitleDelayChange
      ? (delay: number) => deps.onSecondarySubtitleDelayChange!(delay)
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
    broadcastSecondarySubtitle: (text: string) => {
      if (deps.onSecondarySubtitleChange) {
        deps.onSecondarySubtitleChange(text);
        return;
      }
      deps.broadcastToOverlayWindows('secondary-subtitle:set', text);
    },
    updateCurrentMediaPath: (path: string) => {
      resetSubtitleDeduplication();
      deps.updateCurrentMediaPath(path);
    },
    restoreMpvSubVisibility: () => deps.restoreMpvSubVisibility(),
    resetSubtitleSidebarEmbeddedLayout: () => deps.resetSubtitleSidebarEmbeddedLayout?.(),
    getCurrentAnilistMediaKey: () => deps.getCurrentAnilistMediaKey(),
    resetAnilistMediaTracking: (mediaKey: string | null) =>
      deps.resetAnilistMediaTracking(mediaKey),
    maybeProbeAnilistDuration: (mediaKey: string) => deps.maybeProbeAnilistDuration(mediaKey),
    ensureAnilistMediaGuess: (mediaKey: string) => deps.ensureAnilistMediaGuess(mediaKey),
    syncImmersionMediaState: () => deps.syncImmersionMediaState(),
    signalAutoplayReadyIfWarm: (path: string) => deps.signalAutoplayReadyIfWarm?.(path),
    markJellyfinRemotePlaybackLoaded: (path: string) =>
      deps.markJellyfinRemotePlaybackLoaded?.(path),
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
      deps.recordAnilistMediaDuration?.(durationSec);
    },
    reportJellyfinRemoteProgress: (forceImmediate: boolean) =>
      deps.reportJellyfinRemoteProgress(forceImmediate),
    consumeExplicitSeek: deps.consumeExplicitSeek,
    onTimePosUpdate: (time: number) => {
      // Timing history is a viewing log: after a real backward seek, a rewatched
      // canonical line should enter it again. Immersion stats keep their
      // once-per-media deduplication and are not reset here.
      if (
        Number.isFinite(time) &&
        lastTimePosForTimingReset !== null &&
        time <= lastTimePosForTimingReset - SEEK_LIKE_TIME_DELTA_SECONDS
      ) {
        recordedTimingCanonicalKeys.clear();
      }
      if (Number.isFinite(time)) {
        lastTimePosForTimingReset = time;
      }
      deps.onTimePosUpdate?.(time);
    },
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
