import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOnWillQuitCleanupDepsHandler } from './app-lifecycle-main-cleanup';
import { createOnWillQuitCleanupHandler } from './app-lifecycle-actions';

test('cleanup deps builder returns handlers that guard optional runtime objects', () => {
  const calls: string[] = [];
  let reconnectTimer: ReturnType<typeof setTimeout> | null = setTimeout(() => {}, 60_000);
  let immersionTracker: { destroy: () => void } | null = {
    destroy: () => calls.push('destroy-immersion'),
  };

  const depsFactory = createBuildOnWillQuitCleanupDepsHandler({
    destroyTray: () => calls.push('destroy-tray'),
    stopConfigHotReload: () => calls.push('stop-config'),
    restorePreviousSecondarySubVisibility: () => calls.push('restore-sub'),
    restoreMpvSubVisibilityForInvisibleOverlay: () => calls.push('restore-mpv-sub'),
    unregisterAllGlobalShortcuts: () => calls.push('unregister-shortcuts'),
    stopSubtitleWebsocket: () => calls.push('stop-ws'),
    stopTexthookerService: () => calls.push('stop-texthooker'),

    getYomitanParserWindow: () => ({
      isDestroyed: () => false,
      destroy: () => calls.push('destroy-yomitan-window'),
    }),
    clearYomitanParserState: () => calls.push('clear-yomitan-state'),

    getWindowTracker: () => ({ stop: () => calls.push('stop-tracker') }),
    flushMpvLog: () => calls.push('flush-mpv-log'),
    getMpvSocket: () => ({ destroy: () => calls.push('destroy-socket') }),
    getReconnectTimer: () => reconnectTimer,
    clearReconnectTimerRef: () => {
      reconnectTimer = null;
      calls.push('clear-reconnect-ref');
    },

    getSubtitleTimingTracker: () => ({ destroy: () => calls.push('destroy-subtitle-tracker') }),
    getImmersionTracker: () => immersionTracker,
    clearImmersionTracker: () => {
      immersionTracker = null;
      calls.push('clear-immersion-ref');
    },
    getAnkiIntegration: () => ({ destroy: () => calls.push('destroy-anki') }),

    getAnilistSetupWindow: () => ({ destroy: () => calls.push('destroy-anilist-window') }),
    clearAnilistSetupWindow: () => calls.push('clear-anilist-window'),
    getJellyfinSetupWindow: () => ({ destroy: () => calls.push('destroy-jellyfin-window') }),
    clearJellyfinSetupWindow: () => calls.push('clear-jellyfin-window'),

    stopJellyfinRemoteSession: () => calls.push('stop-jellyfin-remote'),
    stopDiscordPresenceService: () => calls.push('stop-discord-presence'),
  });

  const cleanup = createOnWillQuitCleanupHandler(depsFactory());
  cleanup();

  assert.ok(calls.includes('destroy-tray'));
  assert.ok(calls.includes('destroy-yomitan-window'));
  assert.ok(calls.includes('flush-mpv-log'));
  assert.ok(calls.includes('destroy-socket'));
  assert.ok(calls.includes('clear-reconnect-ref'));
  assert.ok(calls.includes('destroy-immersion'));
  assert.ok(calls.includes('clear-immersion-ref'));
  assert.ok(calls.includes('stop-jellyfin-remote'));
  assert.ok(calls.includes('stop-discord-presence'));
  assert.equal(reconnectTimer, null);
  assert.equal(immersionTracker, null);
});

test('cleanup deps builder skips destroyed yomitan window', () => {
  const calls: string[] = [];
  const depsFactory = createBuildOnWillQuitCleanupDepsHandler({
    destroyTray: () => {},
    stopConfigHotReload: () => {},
    restorePreviousSecondarySubVisibility: () => {},
    restoreMpvSubVisibilityForInvisibleOverlay: () => {},
    unregisterAllGlobalShortcuts: () => {},
    stopSubtitleWebsocket: () => {},
    stopTexthookerService: () => {},
    getYomitanParserWindow: () => ({
      isDestroyed: () => true,
      destroy: () => calls.push('destroy-yomitan-window'),
    }),
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
    stopJellyfinRemoteSession: () => {},
    stopDiscordPresenceService: () => {},
  });

  const cleanup = createOnWillQuitCleanupHandler(depsFactory());
  cleanup();

  assert.deepEqual(calls, []);
});
