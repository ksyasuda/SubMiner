// Narrow structural types used by cleanup assembly.
type Destroyable = {
  destroy: () => void;
};

type DestroyableWindow = Destroyable & {
  isDestroyed: () => boolean;
};

type Stoppable = {
  stop: () => void;
};

type SocketLike = {
  destroy: () => void;
};

type TimerLike = ReturnType<typeof setTimeout>;

export function createBuildOnWillQuitCleanupDepsHandler(deps: {
  destroyTray: () => void;
  stopConfigHotReload: () => void;
  restorePreviousSecondarySubVisibility: () => void;
  unregisterAllGlobalShortcuts: () => void;
  stopSubtitleWebsocket: () => void;
  stopTexthookerService: () => void;

  getYomitanParserWindow: () => DestroyableWindow | null;
  clearYomitanParserState: () => void;

  getWindowTracker: () => Stoppable | null;
  flushMpvLog: () => void;
  getMpvSocket: () => SocketLike | null;
  getReconnectTimer: () => TimerLike | null;
  clearReconnectTimerRef: () => void;

  getSubtitleTimingTracker: () => Destroyable | null;
  getImmersionTracker: () => Destroyable | null;
  clearImmersionTracker: () => void;
  getAnkiIntegration: () => Destroyable | null;

  getAnilistSetupWindow: () => Destroyable | null;
  clearAnilistSetupWindow: () => void;
  getJellyfinSetupWindow: () => Destroyable | null;
  clearJellyfinSetupWindow: () => void;

  stopJellyfinRemoteSession: () => void;
}) {
  return () => ({
    destroyTray: () => deps.destroyTray(),
    stopConfigHotReload: () => deps.stopConfigHotReload(),
    restorePreviousSecondarySubVisibility: () => deps.restorePreviousSecondarySubVisibility(),
    unregisterAllGlobalShortcuts: () => deps.unregisterAllGlobalShortcuts(),
    stopSubtitleWebsocket: () => deps.stopSubtitleWebsocket(),
    stopTexthookerService: () => deps.stopTexthookerService(),
    destroyYomitanParserWindow: () => {
      const window = deps.getYomitanParserWindow();
      if (!window) return;
      if (window.isDestroyed()) return;
      window.destroy();
    },
    clearYomitanParserState: () => deps.clearYomitanParserState(),
    stopWindowTracker: () => {
      const tracker = deps.getWindowTracker();
      tracker?.stop();
    },
    flushMpvLog: () => deps.flushMpvLog(),
    destroyMpvSocket: () => {
      const socket = deps.getMpvSocket();
      socket?.destroy();
    },
    clearReconnectTimer: () => {
      const timer = deps.getReconnectTimer();
      if (!timer) return;
      clearTimeout(timer);
      deps.clearReconnectTimerRef();
    },
    destroySubtitleTimingTracker: () => {
      deps.getSubtitleTimingTracker()?.destroy();
    },
    destroyImmersionTracker: () => {
      const tracker = deps.getImmersionTracker();
      if (!tracker) return;
      tracker.destroy();
      deps.clearImmersionTracker();
    },
    destroyAnkiIntegration: () => {
      deps.getAnkiIntegration()?.destroy();
    },
    destroyAnilistSetupWindow: () => {
      deps.getAnilistSetupWindow()?.destroy();
    },
    clearAnilistSetupWindow: () => deps.clearAnilistSetupWindow(),
    destroyJellyfinSetupWindow: () => {
      deps.getJellyfinSetupWindow()?.destroy();
    },
    clearJellyfinSetupWindow: () => deps.clearJellyfinSetupWindow(),
    stopJellyfinRemoteSession: () => deps.stopJellyfinRemoteSession(),
  });
}
