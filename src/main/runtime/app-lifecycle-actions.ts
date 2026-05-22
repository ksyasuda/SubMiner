export function createOnWillQuitCleanupHandler(deps: {
  destroyTray: () => void;
  stopConfigHotReload: () => void;
  restorePreviousSecondarySubVisibility: () => void;
  restoreMpvSubVisibility: () => void;
  unregisterAllGlobalShortcuts: () => void;
  stopSubtitleWebsocket: () => void;
  stopTexthookerService: () => void;
  clearWindowsVisibleOverlayForegroundPollLoop: () => void;
  clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => void;
  destroyMainOverlayWindow: () => void;
  destroyModalOverlayWindow: () => void;
  destroyYomitanParserWindow: () => void;
  clearYomitanParserState: () => void;
  stopWindowTracker: () => void;
  flushMpvLog: () => void;
  destroyMpvSocket: () => void;
  clearReconnectTimer: () => void;
  destroySubtitleTimingTracker: () => void;
  destroyImmersionTracker: () => void;
  destroyAnkiIntegration: () => void;
  destroyAnilistSetupWindow: () => void;
  clearAnilistSetupWindow: () => void;
  destroyJellyfinSetupWindow: () => void;
  clearJellyfinSetupWindow: () => void;
  destroyFirstRunSetupWindow: () => void;
  clearFirstRunSetupWindow: () => void;
  destroyYomitanSettingsWindow: () => void;
  clearYomitanSettingsWindow: () => void;
  stopJellyfinRemoteSession: () => void;
  cleanupYoutubeSubtitleTempDirs: () => void;
  cleanupJellyfinSubtitleCache: () => void;
  stopDiscordPresenceService: () => void;
}) {
  return (): void => {
    deps.destroyTray();
    deps.stopConfigHotReload();
    deps.restorePreviousSecondarySubVisibility();
    deps.restoreMpvSubVisibility();
    deps.unregisterAllGlobalShortcuts();
    deps.stopSubtitleWebsocket();
    deps.stopTexthookerService();
    deps.clearWindowsVisibleOverlayForegroundPollLoop();
    deps.clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    deps.destroyMainOverlayWindow();
    deps.destroyModalOverlayWindow();
    deps.destroyYomitanParserWindow();
    deps.clearYomitanParserState();
    deps.stopWindowTracker();
    deps.flushMpvLog();
    deps.destroyMpvSocket();
    deps.clearReconnectTimer();
    deps.destroySubtitleTimingTracker();
    deps.destroyImmersionTracker();
    deps.destroyAnkiIntegration();
    deps.destroyAnilistSetupWindow();
    deps.clearAnilistSetupWindow();
    deps.destroyJellyfinSetupWindow();
    deps.clearJellyfinSetupWindow();
    deps.destroyFirstRunSetupWindow();
    deps.clearFirstRunSetupWindow();
    deps.destroyYomitanSettingsWindow();
    deps.clearYomitanSettingsWindow();
    deps.stopJellyfinRemoteSession();
    deps.cleanupYoutubeSubtitleTempDirs();
    deps.cleanupJellyfinSubtitleCache();
    deps.stopDiscordPresenceService();
  };
}

export function createShouldRestoreWindowsOnActivateHandler(deps: {
  isOverlayRuntimeInitialized: () => boolean;
  getAllWindowCount: () => number;
}) {
  return (): boolean => deps.isOverlayRuntimeInitialized() && deps.getAllWindowCount() === 0;
}

export function createRestoreWindowsOnActivateHandler(deps: {
  createMainWindow: () => void;
  updateVisibleOverlayVisibility: () => void;
  syncOverlayMpvSubtitleSuppression: () => void;
}) {
  return (): void => {
    deps.createMainWindow();
    deps.updateVisibleOverlayVisibility();
    deps.syncOverlayMpvSubtitleSuppression();
  };
}
