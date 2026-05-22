import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createOnWillQuitCleanupHandler,
  createRestoreWindowsOnActivateHandler,
  createShouldRestoreWindowsOnActivateHandler,
} from './app-lifecycle-actions';

test('on will quit cleanup handler runs all cleanup steps', () => {
  const calls: string[] = [];
  const cleanup = createOnWillQuitCleanupHandler({
    destroyTray: () => calls.push('destroy-tray'),
    stopConfigHotReload: () => calls.push('stop-config'),
    restorePreviousSecondarySubVisibility: () => calls.push('restore-sub'),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    unregisterAllGlobalShortcuts: () => calls.push('unregister-shortcuts'),
    stopSubtitleWebsocket: () => calls.push('stop-ws'),
    stopTexthookerService: () => calls.push('stop-texthooker'),
    clearWindowsVisibleOverlayForegroundPollLoop: () =>
      calls.push('clear-windows-visible-overlay-poll'),
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: () =>
      calls.push('clear-linux-mpv-fullscreen-overlay-refresh-timeouts'),
    destroyMainOverlayWindow: () => calls.push('destroy-main-overlay-window'),
    destroyModalOverlayWindow: () => calls.push('destroy-modal-overlay-window'),
    destroyYomitanParserWindow: () => calls.push('destroy-yomitan-window'),
    clearYomitanParserState: () => calls.push('clear-yomitan-state'),
    stopWindowTracker: () => calls.push('stop-tracker'),
    flushMpvLog: () => calls.push('flush-mpv-log'),
    destroyMpvSocket: () => calls.push('destroy-socket'),
    clearReconnectTimer: () => calls.push('clear-reconnect'),
    destroySubtitleTimingTracker: () => calls.push('destroy-subtitle-tracker'),
    destroyImmersionTracker: () => calls.push('destroy-immersion'),
    destroyAnkiIntegration: () => calls.push('destroy-anki'),
    destroyAnilistSetupWindow: () => calls.push('destroy-anilist-window'),
    clearAnilistSetupWindow: () => calls.push('clear-anilist-window'),
    destroyJellyfinSetupWindow: () => calls.push('destroy-jellyfin-window'),
    clearJellyfinSetupWindow: () => calls.push('clear-jellyfin-window'),
    destroyFirstRunSetupWindow: () => calls.push('destroy-first-run-window'),
    clearFirstRunSetupWindow: () => calls.push('clear-first-run-window'),
    destroyYomitanSettingsWindow: () => calls.push('destroy-yomitan-settings-window'),
    clearYomitanSettingsWindow: () => calls.push('clear-yomitan-settings-window'),
    stopJellyfinRemoteSession: () => calls.push('stop-jellyfin-remote'),
    cleanupYoutubeSubtitleTempDirs: () => calls.push('cleanup-youtube-subtitles'),
    cleanupJellyfinSubtitleCache: () => calls.push('cleanup-jellyfin-subtitles'),
    stopDiscordPresenceService: () => calls.push('stop-discord-presence'),
  });

  cleanup();
  assert.equal(calls.length, 31);
  assert.equal(calls[0], 'destroy-tray');
  assert.equal(calls[calls.length - 1], 'stop-discord-presence');
  assert.ok(calls.includes('cleanup-jellyfin-subtitles'));
  assert.ok(calls.includes('clear-windows-visible-overlay-poll'));
  assert.ok(calls.includes('clear-linux-mpv-fullscreen-overlay-refresh-timeouts'));
  assert.ok(calls.includes('cleanup-youtube-subtitles'));
  assert.ok(calls.indexOf('flush-mpv-log') < calls.indexOf('destroy-socket'));
});

test('on will quit cleanup handler cleans jellyfin subtitle cache when stopping remote session fails', () => {
  const calls: string[] = [];
  const cleanup = createOnWillQuitCleanupHandler({
    destroyTray: () => {},
    stopConfigHotReload: () => {},
    restorePreviousSecondarySubVisibility: () => {},
    restoreMpvSubVisibility: () => {},
    unregisterAllGlobalShortcuts: () => {},
    stopSubtitleWebsocket: () => {},
    stopTexthookerService: () => {},
    clearWindowsVisibleOverlayForegroundPollLoop: () => {},
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => {},
    destroyMainOverlayWindow: () => {},
    destroyModalOverlayWindow: () => {},
    destroyYomitanParserWindow: () => {},
    clearYomitanParserState: () => {},
    stopWindowTracker: () => {},
    flushMpvLog: () => {},
    destroyMpvSocket: () => {},
    clearReconnectTimer: () => {},
    destroySubtitleTimingTracker: () => {},
    destroyImmersionTracker: () => {},
    destroyAnkiIntegration: () => {},
    destroyAnilistSetupWindow: () => {},
    clearAnilistSetupWindow: () => {},
    destroyJellyfinSetupWindow: () => {},
    clearJellyfinSetupWindow: () => {},
    destroyFirstRunSetupWindow: () => {},
    clearFirstRunSetupWindow: () => {},
    destroyYomitanSettingsWindow: () => {},
    clearYomitanSettingsWindow: () => {},
    stopJellyfinRemoteSession: () => {
      calls.push('stop-jellyfin-remote');
      throw new Error('stop failed');
    },
    cleanupYoutubeSubtitleTempDirs: () => calls.push('cleanup-youtube-subtitles'),
    cleanupJellyfinSubtitleCache: () => calls.push('cleanup-jellyfin-subtitles'),
    stopDiscordPresenceService: () => calls.push('stop-discord-presence'),
  });

  assert.throws(() => cleanup(), /stop failed/);
  assert.deepEqual(calls, ['stop-jellyfin-remote', 'cleanup-jellyfin-subtitles']);
});

test('should restore windows on activate requires initialized runtime and no windows', () => {
  let initialized = false;
  let windowCount = 1;
  const shouldRestore = createShouldRestoreWindowsOnActivateHandler({
    isOverlayRuntimeInitialized: () => initialized,
    getAllWindowCount: () => windowCount,
  });

  assert.equal(shouldRestore(), false);
  initialized = true;
  assert.equal(shouldRestore(), false);
  windowCount = 0;
  assert.equal(shouldRestore(), true);
});

test('restore windows on activate recreates windows then syncs visibility', () => {
  const calls: string[] = [];
  const restore = createRestoreWindowsOnActivateHandler({
    createMainWindow: () => calls.push('main'),
    updateVisibleOverlayVisibility: () => calls.push('visible-sync'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('mpv-sync'),
  });

  restore();
  assert.deepEqual(calls, ['main', 'visible-sync', 'mpv-sync']);
});
