import assert from 'node:assert/strict';
import test from 'node:test';
import { composeStartupLifecycleHandlers } from './startup-lifecycle-composer';

test('composeStartupLifecycleHandlers returns callable startup lifecycle handlers', () => {
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
      restoreMpvSubVisibilityForInvisibleOverlay: () => {},
      unregisterAllGlobalShortcuts: () => {},
      stopSubtitleWebsocket: () => {},
      stopTexthookerService: () => {},
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
      stopJellyfinRemoteSession: async () => {},
      stopDiscordPresenceService: () => {},
    },
    shouldRestoreWindowsOnActivateMainDeps: {
      isOverlayRuntimeInitialized: () => false,
      getAllWindowCount: () => 0,
    },
    restoreWindowsOnActivateMainDeps: {
      createMainWindow: () => {},
      updateVisibleOverlayVisibility: () => {},
      syncOverlayMpvSubtitleSuppression: () => {},
    },
  });

  assert.equal(typeof composed.registerProtocolUrlHandlers, 'function');
  assert.equal(typeof composed.onWillQuitCleanup, 'function');
  assert.equal(typeof composed.shouldRestoreWindowsOnActivate, 'function');
  assert.equal(typeof composed.restoreWindowsOnActivate, 'function');
});
