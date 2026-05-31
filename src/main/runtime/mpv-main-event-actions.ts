import type { SubtitleData } from '../../types';

type AnilistPostWatchRunOptions = {
  watchedSeconds?: number;
};

const SEEK_LIKE_TIME_DELTA_SECONDS = 2.5;

function isSeekLikeTimeChange(previousTime: number | null, nextTime: number): boolean {
  if (previousTime === null || !Number.isFinite(previousTime) || !Number.isFinite(nextTime)) {
    return false;
  }
  return Math.abs(nextTime - previousTime) >= SEEK_LIKE_TIME_DELTA_SECONDS;
}

export function createHandleMpvSubtitleChangeHandler(deps: {
  setCurrentSubText: (text: string) => void;
  getImmediateSubtitlePayload?: (text: string) => SubtitleData | null;
  emitImmediateSubtitle?: (payload: SubtitleData) => void;
  broadcastSubtitle: (payload: SubtitleData) => void;
  onSubtitleChange: (text: string) => void;
  refreshDiscordPresence: () => void;
}) {
  return ({ text }: { text: string }): void => {
    deps.setCurrentSubText(text);
    const immediatePayload = deps.getImmediateSubtitlePayload?.(text) ?? null;
    if (immediatePayload) {
      deps.onSubtitleChange(text);
      (deps.emitImmediateSubtitle ?? deps.broadcastSubtitle)(immediatePayload);
    } else {
      if (!text.trim()) {
        deps.broadcastSubtitle({
          text,
          tokens: null,
        });
      }
      deps.onSubtitleChange(text);
    }
    deps.refreshDiscordPresence();
  };
}

export function createHandleMpvSubtitleAssChangeHandler(deps: {
  setCurrentSubAssText: (text: string) => void;
  broadcastSubtitleAss: (text: string) => void;
}) {
  return ({ text }: { text: string }): void => {
    deps.setCurrentSubAssText(text);
    deps.broadcastSubtitleAss(text);
  };
}

export function createHandleMpvSecondarySubtitleChangeHandler(deps: {
  broadcastSecondarySubtitle: (text: string) => void;
}) {
  return ({ text }: { text: string }): void => {
    deps.broadcastSecondarySubtitle(text);
  };
}

export function createHandleMpvMediaPathChangeHandler(deps: {
  updateCurrentMediaPath: (path: string) => void;
  reportJellyfinRemoteStopped: () => void;
  restoreMpvSubVisibility: () => void;
  resetSubtitleSidebarEmbeddedLayout: () => void;
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  maybeProbeAnilistDuration: (mediaKey: string) => void;
  ensureAnilistMediaGuess: (mediaKey: string) => void;
  syncImmersionMediaState: () => void;
  scheduleCharacterDictionarySync?: () => void;
  signalAutoplayReadyIfWarm?: (path: string) => void;
  markJellyfinRemotePlaybackLoaded?: (path: string) => void;
  flushPlaybackPositionOnMediaPathClear?: (mediaPath: string) => void;
  refreshDiscordPresence: () => void;
}) {
  return ({ path }: { path: string | null }): void => {
    const normalizedPath = typeof path === 'string' ? path : '';
    if (!normalizedPath) {
      deps.flushPlaybackPositionOnMediaPathClear?.(normalizedPath);
    }
    deps.updateCurrentMediaPath(normalizedPath);
    deps.resetSubtitleSidebarEmbeddedLayout();
    if (!normalizedPath) {
      deps.reportJellyfinRemoteStopped();
      deps.restoreMpvSubVisibility();
    }
    const mediaKey = deps.getCurrentAnilistMediaKey();
    deps.resetAnilistMediaTracking(mediaKey);
    if (mediaKey) {
      deps.maybeProbeAnilistDuration(mediaKey);
      deps.ensureAnilistMediaGuess(mediaKey);
    }
    deps.syncImmersionMediaState();
    if (normalizedPath.trim().length > 0) {
      deps.markJellyfinRemotePlaybackLoaded?.(normalizedPath);
      deps.scheduleCharacterDictionarySync?.();
      deps.signalAutoplayReadyIfWarm?.(normalizedPath);
    }
    deps.refreshDiscordPresence();
  };
}

export function createHandleMpvMediaTitleChangeHandler(deps: {
  updateCurrentMediaTitle: (title: string) => void;
  resetAnilistMediaGuessState: () => void;
  notifyImmersionTitleUpdate: (title: string) => void;
  syncImmersionMediaState: () => void;
  refreshDiscordPresence: () => void;
}) {
  return ({ title }: { title: string | null }): void => {
    const normalizedTitle = typeof title === 'string' ? title : '';
    deps.updateCurrentMediaTitle(normalizedTitle);
    deps.resetAnilistMediaGuessState();
    deps.notifyImmersionTitleUpdate(normalizedTitle);
    deps.syncImmersionMediaState();
    deps.refreshDiscordPresence();
  };
}

export function createHandleMpvTimePosChangeHandler(deps: {
  recordPlaybackPosition: (time: number) => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  refreshDiscordPresence: () => void;
  maybeRunAnilistPostWatchUpdate?: (options?: AnilistPostWatchRunOptions) => Promise<void>;
  logError?: (message: string, error: unknown) => void;
  onTimePosUpdate?: (time: number) => void;
}) {
  let lastObservedTime: number | null = null;

  return ({ time }: { time: number }): void => {
    const forceImmediate = isSeekLikeTimeChange(lastObservedTime, time);
    if (Number.isFinite(time)) {
      lastObservedTime = time;
    }
    deps.recordPlaybackPosition(time);
    deps.reportJellyfinRemoteProgress(forceImmediate);
    deps.refreshDiscordPresence();
    void deps.maybeRunAnilistPostWatchUpdate?.({ watchedSeconds: time }).catch((error) => {
      deps.logError?.('AniList post-watch update failed unexpectedly', error);
    });
    deps.onTimePosUpdate?.(time);
  };
}

export function createHandleMpvPauseChangeHandler(deps: {
  recordPauseState: (paused: boolean) => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  refreshDiscordPresence: () => void;
}) {
  return ({ paused }: { paused: boolean }): void => {
    deps.recordPauseState(paused);
    deps.reportJellyfinRemoteProgress(true);
    deps.refreshDiscordPresence();
  };
}

export function createHandleMpvSubtitleMetricsChangeHandler(deps: {
  updateSubtitleRenderMetrics: (patch: Record<string, unknown>) => void;
}) {
  return ({ patch }: { patch: Record<string, unknown> }): void => {
    deps.updateSubtitleRenderMetrics(patch);
  };
}

export function createHandleMpvSecondarySubtitleVisibilityHandler(deps: {
  setPreviousSecondarySubVisibility: (visible: boolean) => void;
}) {
  return ({ visible }: { visible: boolean }): void => {
    deps.setPreviousSecondarySubVisibility(visible);
  };
}
