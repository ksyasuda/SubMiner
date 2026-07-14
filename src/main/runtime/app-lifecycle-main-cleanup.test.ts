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
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    isAppReady: () => true,
    unregisterAllGlobalShortcuts: () => calls.push('unregister-shortcuts'),
    stopSubtitleWebsocket: () => calls.push('stop-ws'),
    stopTexthookerService: () => calls.push('stop-texthooker'),
    stopSyncAutoScheduler: () => {
      calls.push('stop-sync-auto-scheduler');
    },
    clearWindowsVisibleOverlayForegroundPollLoop: () =>
      calls.push('clear-windows-visible-overlay-foreground-poll-loop'),
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: () =>
      calls.push('clear-linux-mpv-fullscreen-overlay-refresh-timeouts'),
    getMainOverlayWindow: () => ({
      isDestroyed: () => false,
      destroy: () => calls.push('destroy-main-overlay-window'),
    }),
    clearMainOverlayWindow: () => calls.push('clear-main-overlay-window'),
    getModalOverlayWindow: () => ({
      isDestroyed: () => false,
      destroy: () => calls.push('destroy-modal-overlay-window'),
    }),
    clearModalOverlayWindow: () => calls.push('clear-modal-overlay-window'),

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
    getFirstRunSetupWindow: () => ({ destroy: () => calls.push('destroy-first-run-window') }),
    clearFirstRunSetupWindow: () => calls.push('clear-first-run-window'),
    getYomitanSettingsWindow: () => ({
      destroy: () => calls.push('destroy-yomitan-settings-window'),
    }),
    clearYomitanSettingsWindow: () => calls.push('clear-yomitan-settings-window'),

    stopJellyfinRemoteSession: () => calls.push('stop-jellyfin-remote'),
    cleanupYoutubeSubtitleTempDirs: () => calls.push('cleanup-youtube-subtitles'),
    cleanupYoutubeMediaCache: () => calls.push('cleanup-youtube-media'),
    cleanupJellyfinSubtitleCache: () => calls.push('cleanup-jellyfin-subtitles'),
    stopDiscordPresenceService: () => calls.push('stop-discord-presence'),
  });

  const cleanup = createOnWillQuitCleanupHandler(depsFactory());
  cleanup();

  assert.ok(calls.includes('destroy-tray'));
  assert.ok(calls.includes('destroy-main-overlay-window'));
  assert.ok(calls.includes('clear-main-overlay-window'));
  assert.ok(calls.includes('destroy-modal-overlay-window'));
  assert.ok(calls.includes('clear-modal-overlay-window'));
  assert.ok(calls.includes('destroy-yomitan-window'));
  assert.ok(calls.includes('flush-mpv-log'));
  assert.ok(calls.includes('destroy-socket'));
  assert.ok(calls.includes('clear-reconnect-ref'));
  assert.ok(calls.includes('destroy-immersion'));
  assert.ok(calls.includes('clear-immersion-ref'));
  assert.ok(calls.includes('destroy-first-run-window'));
  assert.ok(calls.includes('destroy-yomitan-settings-window'));
  assert.ok(calls.includes('stop-jellyfin-remote'));
  assert.ok(calls.includes('cleanup-youtube-subtitles'));
  assert.ok(calls.includes('cleanup-youtube-media'));
  assert.ok(calls.includes('cleanup-jellyfin-subtitles'));
  assert.ok(calls.includes('stop-discord-presence'));
  assert.ok(calls.includes('clear-windows-visible-overlay-foreground-poll-loop'));
  assert.ok(calls.includes('clear-linux-mpv-fullscreen-overlay-refresh-timeouts'));
  assert.equal(reconnectTimer, null);
  assert.equal(immersionTracker, null);
});

test('cleanup deps builder skips destroyed yomitan window', () => {
  const calls: string[] = [];
  const depsFactory = createBuildOnWillQuitCleanupDepsHandler({
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
    getMainOverlayWindow: () => ({
      isDestroyed: () => true,
      destroy: () => calls.push('destroy-main-overlay-window'),
    }),
    clearMainOverlayWindow: () => calls.push('clear-main-overlay-window'),
    getModalOverlayWindow: () => ({
      isDestroyed: () => true,
      destroy: () => calls.push('destroy-modal-overlay-window'),
    }),
    clearModalOverlayWindow: () => calls.push('clear-modal-overlay-window'),
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
    getFirstRunSetupWindow: () => null,
    clearFirstRunSetupWindow: () => {},
    getYomitanSettingsWindow: () => null,
    clearYomitanSettingsWindow: () => {},
    stopJellyfinRemoteSession: () => {},
    cleanupYoutubeSubtitleTempDirs: () => {},
    cleanupYoutubeMediaCache: () => {},
    cleanupJellyfinSubtitleCache: () => {},
    stopDiscordPresenceService: () => {},
  });

  const cleanup = createOnWillQuitCleanupHandler(depsFactory());
  cleanup();

  assert.deepEqual(calls, []);
});

test('cleanup deps builder skips global shortcut cleanup before app ready', () => {
  const calls: string[] = [];
  const depsFactory = createBuildOnWillQuitCleanupDepsHandler({
    destroyTray: () => calls.push('destroy-tray'),
    stopConfigHotReload: () => {},
    restorePreviousSecondarySubVisibility: () => {},
    restoreMpvSubVisibility: () => {},
    isAppReady: () => false,
    unregisterAllGlobalShortcuts: () => {
      throw new Error('globalShortcut cannot be used before the app is ready');
    },
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
    stopJellyfinRemoteSession: () => {},
    cleanupYoutubeSubtitleTempDirs: () => {},
    cleanupYoutubeMediaCache: () => {},
    cleanupJellyfinSubtitleCache: () => {},
    stopDiscordPresenceService: () => {},
  });

  const cleanup = createOnWillQuitCleanupHandler(depsFactory());
  cleanup();

  assert.deepEqual(calls, ['destroy-tray']);
});
