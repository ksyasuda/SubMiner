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
    unregisterAllGlobalShortcuts: () => calls.push('unregister-shortcuts'),
    stopSubtitleWebsocket: () => calls.push('stop-ws'),
    stopTexthookerService: () => calls.push('stop-texthooker'),
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
    stopJellyfinRemoteSession: () => calls.push('stop-jellyfin-remote'),
    stopDiscordPresenceService: () => calls.push('stop-discord-presence'),
  });

  cleanup();
  assert.equal(calls.length, 21);
  assert.equal(calls[0], 'destroy-tray');
  assert.equal(calls[calls.length - 1], 'stop-discord-presence');
  assert.ok(calls.indexOf('flush-mpv-log') < calls.indexOf('destroy-socket'));
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
    createInvisibleWindow: () => calls.push('invisible'),
    updateVisibleOverlayVisibility: () => calls.push('visible-sync'),
    updateInvisibleOverlayVisibility: () => calls.push('invisible-sync'),
  });

  restore();
  assert.deepEqual(calls, ['main', 'invisible', 'visible-sync', 'invisible-sync']);
});
