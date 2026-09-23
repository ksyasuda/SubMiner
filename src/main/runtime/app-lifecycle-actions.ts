export function createForceQuitHandler(deps: {
  destroyImmersionTracker: () => void | Promise<void>;
  logError: (error: unknown) => void;
  exit: () => void;
}) {
  return async () => {
    let timeout: ReturnType<typeof setTimeout> | undefined;
    try {
      await Promise.race([
        Promise.resolve().then(() => deps.destroyImmersionTracker()),
        new Promise<never>((_, reject) => {
          timeout = setTimeout(() => reject(new Error('Stats finalization timed out.')), 1_000);
        }),
      ]);
    } catch (error) {
      deps.logError(error);
    } finally {
      clearTimeout(timeout);
      deps.exit();
    }
  };
}

export function createOnWillQuitCleanupHandler(deps: {
  destroyTray: () => void;
  stopConfigHotReload: () => void;
  restorePreviousSecondarySubVisibility: () => void;
  restoreMpvSubVisibility: () => void;
  unregisterAllGlobalShortcuts: () => void;
  stopSubtitleWebsocket: () => void;
  stopTexthookerService: () => void;
  stopSyncAutoScheduler: () => void | Promise<void>;
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
  stopStatsServer: () => Promise<void> | void;
  destroyImmersionTracker: () => void | Promise<void>;
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
  cleanupInternalSubtitleTrackCache: () => void;
  cleanupYoutubeSubtitleTempDirs: () => void;
  cleanupYoutubeMediaCache: () => void;
  cleanupRemoteMediaWindows: () => void;
  cleanupJellyfinSubtitleCache: () => void;
  stopDiscordPresenceService: () => void;
}) {
  return async (): Promise<void> => {
    const cleanupErrors: unknown[] = [];
    deps.destroyTray();
    deps.stopConfigHotReload();
    deps.restorePreviousSecondarySubVisibility();
    deps.restoreMpvSubVisibility();
    deps.unregisterAllGlobalShortcuts();
    deps.stopSubtitleWebsocket();
    deps.stopTexthookerService();
    const stopSyncAutoScheduler = Promise.resolve(deps.stopSyncAutoScheduler()).catch(
      (error: unknown) => {
        cleanupErrors.push(error);
      },
    );
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
    try {
      await deps.stopStatsServer();
    } catch (error) {
      cleanupErrors.push(error);
    }
    try {
      await deps.destroyImmersionTracker();
    } catch (error) {
      cleanupErrors.push(error);
    }
    deps.destroyAnkiIntegration();
    deps.destroyAnilistSetupWindow();
    deps.clearAnilistSetupWindow();
    deps.destroyJellyfinSetupWindow();
    deps.clearJellyfinSetupWindow();
    deps.destroyFirstRunSetupWindow();
    deps.clearFirstRunSetupWindow();
    deps.destroyYomitanSettingsWindow();
    deps.clearYomitanSettingsWindow();
    const runCleanup = (cleanup: () => void): void => {
      try {
        cleanup();
      } catch (error) {
        cleanupErrors.push(error);
      }
    };
    try {
      runCleanup(deps.stopJellyfinRemoteSession);
      runCleanup(deps.cleanupJellyfinSubtitleCache);
      runCleanup(deps.cleanupInternalSubtitleTrackCache);
    } finally {
      runCleanup(deps.cleanupYoutubeSubtitleTempDirs);
      runCleanup(deps.cleanupYoutubeMediaCache);
      runCleanup(deps.cleanupRemoteMediaWindows);
      runCleanup(deps.stopDiscordPresenceService);
      await stopSyncAutoScheduler;
    }
    if (cleanupErrors.length > 0) throw cleanupErrors[0];
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
