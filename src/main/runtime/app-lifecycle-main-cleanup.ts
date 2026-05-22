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
  restoreMpvSubVisibility: () => void;
  isAppReady: () => boolean;
  unregisterAllGlobalShortcuts: () => void;
  stopSubtitleWebsocket: () => void;
  stopTexthookerService: () => void;
  clearWindowsVisibleOverlayForegroundPollLoop: () => void;
  clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => void;
  getMainOverlayWindow: () => DestroyableWindow | null;
  clearMainOverlayWindow: () => void;
  getModalOverlayWindow: () => DestroyableWindow | null;
  clearModalOverlayWindow: () => void;

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
  getFirstRunSetupWindow: () => Destroyable | null;
  clearFirstRunSetupWindow: () => void;
  getYomitanSettingsWindow: () => Destroyable | null;
  clearYomitanSettingsWindow: () => void;

  stopJellyfinRemoteSession: () => void;
  cleanupYoutubeSubtitleTempDirs: () => void;
  cleanupJellyfinSubtitleCache: () => void;
  stopDiscordPresenceService: () => void;
}) {
  return () => ({
    destroyTray: () => deps.destroyTray(),
    stopConfigHotReload: () => deps.stopConfigHotReload(),
    restorePreviousSecondarySubVisibility: () => deps.restorePreviousSecondarySubVisibility(),
    restoreMpvSubVisibility: () => deps.restoreMpvSubVisibility(),
    unregisterAllGlobalShortcuts: () => {
      if (!deps.isAppReady()) return;
      deps.unregisterAllGlobalShortcuts();
    },
    stopSubtitleWebsocket: () => deps.stopSubtitleWebsocket(),
    stopTexthookerService: () => deps.stopTexthookerService(),
    clearWindowsVisibleOverlayForegroundPollLoop: () =>
      deps.clearWindowsVisibleOverlayForegroundPollLoop(),
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: () =>
      deps.clearLinuxMpvFullscreenOverlayRefreshTimeouts(),
    destroyMainOverlayWindow: () => {
      const window = deps.getMainOverlayWindow();
      if (!window) return;
      if (window.isDestroyed()) return;
      window.destroy();
      deps.clearMainOverlayWindow();
    },
    destroyModalOverlayWindow: () => {
      const window = deps.getModalOverlayWindow();
      if (!window) return;
      if (window.isDestroyed()) return;
      window.destroy();
      deps.clearModalOverlayWindow();
    },
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
    destroyFirstRunSetupWindow: () => {
      deps.getFirstRunSetupWindow()?.destroy();
    },
    clearFirstRunSetupWindow: () => deps.clearFirstRunSetupWindow(),
    destroyYomitanSettingsWindow: () => {
      deps.getYomitanSettingsWindow()?.destroy();
    },
    clearYomitanSettingsWindow: () => deps.clearYomitanSettingsWindow(),
    stopJellyfinRemoteSession: () => deps.stopJellyfinRemoteSession(),
    cleanupYoutubeSubtitleTempDirs: () => deps.cleanupYoutubeSubtitleTempDirs(),
    cleanupJellyfinSubtitleCache: () => deps.cleanupJellyfinSubtitleCache(),
    stopDiscordPresenceService: () => deps.stopDiscordPresenceService(),
  });
}
