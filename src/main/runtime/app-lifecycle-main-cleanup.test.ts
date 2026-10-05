import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOnWillQuitCleanupDepsHandler } from './app-lifecycle-main-cleanup';
import { createOnWillQuitCleanupHandler } from './app-lifecycle-actions';

type BuilderDeps = Parameters<typeof createBuildOnWillQuitCleanupDepsHandler>[0];

// Defaults model an app that is ready but has no live runtime objects.
function buildCleanupDeps(overrides: Partial<BuilderDeps> = {}) {
  const noop = () => {};
  const none = () => null;
  return createBuildOnWillQuitCleanupDepsHandler({
    destroyTray: noop,
    stopConfigHotReload: noop,
    restorePreviousSecondarySubVisibility: noop,
    restoreMpvSubVisibility: noop,
    isAppReady: () => true,
    unregisterAllGlobalShortcuts: noop,
    stopSubtitleWebsocket: noop,
    stopTexthookerService: noop,
    stopSyncAutoScheduler: noop,
    clearWindowsVisibleOverlayForegroundPollLoop: noop,
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: noop,
    getMainOverlayWindow: none,
    clearMainOverlayWindow: noop,
    getModalOverlayWindow: none,
    clearModalOverlayWindow: noop,
    getYomitanParserWindow: none,
    clearYomitanParserState: noop,
    getWindowTracker: none,
    flushMpvLog: noop,
    getMpvSocket: none,
    getReconnectTimer: none,
    clearReconnectTimerRef: noop,
    getSubtitleTimingTracker: none,
    getImmersionTracker: none,
    stopStatsServer: noop,
    clearImmersionTracker: noop,
    getAnkiIntegration: none,
    getAnilistSetupWindow: none,
    clearAnilistSetupWindow: noop,
    getJellyfinSetupWindow: none,
    clearJellyfinSetupWindow: noop,
    getFirstRunSetupWindow: none,
    clearFirstRunSetupWindow: noop,
    getYomitanSettingsWindow: none,
    clearYomitanSettingsWindow: noop,
    stopJellyfinRemoteSession: noop,
    cleanupInternalSubtitleTrackCache: noop,
    cleanupYoutubeSubtitleTempDirs: noop,
    cleanupYoutubeMediaCache: noop,
    cleanupRemoteMediaWindows: noop,
    cleanupJellyfinSubtitleCache: noop,
    stopDiscordPresenceService: noop,
    ...overrides,
  })();
}

function fakeWindow(calls: string[], destroyed: boolean) {
  return {
    isDestroyed: () => destroyed,
    destroy: () => calls.push('destroy'),
  };
}

test('cleanup deps builder runs the full quit cleanup with no runtime objects', async () => {
  await createOnWillQuitCleanupHandler(buildCleanupDeps())();
});

for (const destroyed of [false, true]) {
  test(`cleanup deps builder ${destroyed ? 'skips' : 'destroys and clears'} ${destroyed ? 'already destroyed' : 'live'} overlay windows`, () => {
    const calls: string[] = [];
    const deps = buildCleanupDeps({
      getMainOverlayWindow: () => fakeWindow(calls, destroyed),
      clearMainOverlayWindow: () => calls.push('clear'),
      getModalOverlayWindow: () => fakeWindow(calls, destroyed),
      clearModalOverlayWindow: () => calls.push('clear'),
      getYomitanParserWindow: () => fakeWindow(calls, destroyed),
    });

    deps.destroyMainOverlayWindow();
    deps.destroyModalOverlayWindow();
    deps.destroyYomitanParserWindow();

    assert.deepEqual(calls, destroyed ? [] : ['destroy', 'clear', 'destroy', 'clear', 'destroy']);
  });
}

test('cleanup deps builder skips global shortcut cleanup before app ready', () => {
  const deps = buildCleanupDeps({
    isAppReady: () => false,
    unregisterAllGlobalShortcuts: () => {
      throw new Error('globalShortcut cannot be used before the app is ready');
    },
  });

  deps.unregisterAllGlobalShortcuts();
});

test('cleanup deps builder clears the reconnect timer and its reference', () => {
  let reconnectTimer: ReturnType<typeof setTimeout> | null = setTimeout(() => {
    assert.fail('reconnect timer should have been cleared');
  }, 0);
  const deps = buildCleanupDeps({
    getReconnectTimer: () => reconnectTimer,
    clearReconnectTimerRef: () => {
      reconnectTimer = null;
    },
  });

  deps.clearReconnectTimer();

  assert.equal(reconnectTimer, null);
});

test('cleanup deps builder clears the immersion tracker only after destroy settles', async () => {
  const calls: string[] = [];
  let immersionTracker: { destroy: () => Promise<void> } | null = {
    destroy: async () => {
      await Promise.resolve();
      calls.push('destroy');
    },
  };
  const deps = buildCleanupDeps({
    getImmersionTracker: () => immersionTracker,
    clearImmersionTracker: () => {
      immersionTracker = null;
      calls.push('clear');
    },
  });

  await deps.destroyImmersionTracker();

  assert.deepEqual(calls, ['destroy', 'clear']);
  assert.equal(immersionTracker, null);
});
