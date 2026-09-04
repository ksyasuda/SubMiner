import assert from 'node:assert/strict';
import test from 'node:test';
import { composeStartupLifecycleHandlers } from './startup-lifecycle-composer';

test('composeStartupLifecycleHandlers returns callable startup lifecycle handlers', () => {
  const calls: string[] = [];
  const composed = composeStartupLifecycleHandlers({
    registerProtocolUrlHandlersMainDeps: {
      registerOpenUrl: () => {},
      registerSecondInstance: () => {},
      handleAnilistSetupProtocolUrl: () => false,
      findAnilistSetupDeepLinkArgvUrl: () => null,
      logUnhandledOpenUrl: () => {},
      logUnhandledSecondInstanceUrl: () => {},
    },
    onWillQuitCleanupMainDeps: {
      destroyTray: () => {},
      stopConfigHotReload: () => {},
      restorePreviousSecondarySubVisibility: () => {},
      restoreMpvSubVisibility: () => {},
      isAppReady: () => true,
      unregisterAllGlobalShortcuts: () => {},
      stopSubtitleWebsocket: () => {},
      stopTexthookerService: () => {},
      stopSyncAutoScheduler: () => {},
      clearWindowsVisibleOverlayForegroundPollLoop: () => {},
      clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => {},
      getMainOverlayWindow: () => null,
      clearMainOverlayWindow: () => {},
      getModalOverlayWindow: () => null,
      clearModalOverlayWindow: () => {},
      getYomitanParserWindow: () => null,
      clearYomitanParserState: () => {},
      getWindowTracker: () => null,
      flushMpvLog: () => {},
      getMpvSocket: () => null,
      getReconnectTimer: () => null,
      clearReconnectTimerRef: () => {},
      getSubtitleTimingTracker: () => null,
      getImmersionTracker: () => null,
      clearImmersionTracker: () => {},
      getAnkiIntegration: () => null,
      getAnilistSetupWindow: () => null,
      clearAnilistSetupWindow: () => {},
      getJellyfinSetupWindow: () => null,
      clearJellyfinSetupWindow: () => {},
      getFirstRunSetupWindow: () => null,
      clearFirstRunSetupWindow: () => {},
      getYomitanSettingsWindow: () => null,
      clearYomitanSettingsWindow: () => {},
      stopJellyfinRemoteSession: async () => {},
      cleanupInternalSubtitleTrackCache: () => {},
      cleanupYoutubeSubtitleTempDirs: () => {},
      cleanupYoutubeMediaCache: () => {},
      cleanupRemoteMediaWindows: () => {},
      cleanupJellyfinSubtitleCache: () => {},
      stopDiscordPresenceService: () => {},
    },
    shouldRestoreWindowsOnActivateMainDeps: {
      isOverlayRuntimeInitialized: () => false,
      getAllWindowCount: () => 0,
    },
    restoreWindowsOnActivateMainDeps: {
      createMainWindow: () => {
        calls.push('createMainWindow');
      },
      updateVisibleOverlayVisibility: () => {},
      syncOverlayMpvSubtitleSuppression: () => {},
    },
  });

  assert.equal(typeof composed.registerProtocolUrlHandlers, 'function');
  assert.equal(typeof composed.onWillQuitCleanup, 'function');
  assert.equal(typeof composed.shouldRestoreWindowsOnActivate, 'function');
  assert.equal(typeof composed.restoreWindowsOnActivate, 'function');

  // shouldRestoreWindowsOnActivate returns false when overlay runtime is not initialized
  assert.equal(composed.shouldRestoreWindowsOnActivate(), false);

  // restoreWindowsOnActivate invokes the injected createMainWindow dep
  composed.restoreWindowsOnActivate();
  assert.deepEqual(calls, ['createMainWindow']);
});
