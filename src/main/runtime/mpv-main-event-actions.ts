import type { SubtitleData } from '../../types';

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
      (deps.emitImmediateSubtitle ?? deps.broadcastSubtitle)(immediatePayload);
    } else {
      deps.broadcastSubtitle({
        text,
        tokens: null,
      });
    }
    deps.onSubtitleChange(text);
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
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  maybeProbeAnilistDuration: (mediaKey: string) => void;
  ensureAnilistMediaGuess: (mediaKey: string) => void;
  syncImmersionMediaState: () => void;
  scheduleCharacterDictionarySync?: () => void;
  signalAutoplayReadyIfWarm?: (path: string) => void;
  refreshDiscordPresence: () => void;
}) {
  return ({ path }: { path: string | null }): void => {
    const normalizedPath = typeof path === 'string' ? path : '';
    deps.updateCurrentMediaPath(normalizedPath);
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
  onTimePosUpdate?: (time: number) => void;
}) {
  return ({ time }: { time: number }): void => {
    deps.recordPlaybackPosition(time);
    deps.reportJellyfinRemoteProgress(false);
    deps.refreshDiscordPresence();
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
