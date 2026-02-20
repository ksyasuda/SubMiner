export function createHandleMpvSubtitleChangeHandler(deps: {
  setCurrentSubText: (text: string) => void;
  broadcastSubtitle: (payload: { text: string; tokens: null }) => void;
  onSubtitleChange: (text: string) => void;
}) {
  return ({ text }: { text: string }): void => {
    deps.setCurrentSubText(text);
    deps.broadcastSubtitle({ text, tokens: null });
    deps.onSubtitleChange(text);
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
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  maybeProbeAnilistDuration: (mediaKey: string) => void;
  ensureAnilistMediaGuess: (mediaKey: string) => void;
  syncImmersionMediaState: () => void;
}) {
  return ({ path }: { path: string }): void => {
    deps.updateCurrentMediaPath(path);
    if (!path) {
      deps.reportJellyfinRemoteStopped();
    }
    const mediaKey = deps.getCurrentAnilistMediaKey();
    deps.resetAnilistMediaTracking(mediaKey);
    if (mediaKey) {
      deps.maybeProbeAnilistDuration(mediaKey);
      deps.ensureAnilistMediaGuess(mediaKey);
    }
    deps.syncImmersionMediaState();
  };
}

export function createHandleMpvMediaTitleChangeHandler(deps: {
  updateCurrentMediaTitle: (title: string) => void;
  resetAnilistMediaGuessState: () => void;
  notifyImmersionTitleUpdate: (title: string) => void;
  syncImmersionMediaState: () => void;
}) {
  return ({ title }: { title: string }): void => {
    deps.updateCurrentMediaTitle(title);
    deps.resetAnilistMediaGuessState();
    deps.notifyImmersionTitleUpdate(title);
    deps.syncImmersionMediaState();
  };
}

export function createHandleMpvTimePosChangeHandler(deps: {
  recordPlaybackPosition: (time: number) => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
}) {
  return ({ time }: { time: number }): void => {
    deps.recordPlaybackPosition(time);
    deps.reportJellyfinRemoteProgress(false);
  };
}

export function createHandleMpvPauseChangeHandler(deps: {
  recordPauseState: (paused: boolean) => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
}) {
  return ({ paused }: { paused: boolean }): void => {
    deps.recordPauseState(paused);
    deps.reportJellyfinRemoteProgress(true);
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
