import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createForceQuitHandler,
  createOnWillQuitCleanupHandler,
  createShouldRestoreWindowsOnActivateHandler,
} from './app-lifecycle-actions';

type CleanupDeps = Parameters<typeof createOnWillQuitCleanupHandler>[0];

// Every cleanup step records its own dep name into `calls` unless overridden.
function makeCleanupDeps(calls: string[], overrides: Partial<CleanupDeps> = {}): CleanupDeps {
  const record = (name: string) => () => {
    calls.push(name);
  };
  return {
    destroyTray: record('destroyTray'),
    stopConfigHotReload: record('stopConfigHotReload'),
    restorePreviousSecondarySubVisibility: record('restorePreviousSecondarySubVisibility'),
    restoreMpvSubVisibility: record('restoreMpvSubVisibility'),
    unregisterAllGlobalShortcuts: record('unregisterAllGlobalShortcuts'),
    stopSubtitleWebsocket: record('stopSubtitleWebsocket'),
    stopTexthookerService: record('stopTexthookerService'),
    stopSyncAutoScheduler: record('stopSyncAutoScheduler'),
    clearWindowsVisibleOverlayForegroundPollLoop: record(
      'clearWindowsVisibleOverlayForegroundPollLoop',
    ),
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: record(
      'clearLinuxMpvFullscreenOverlayRefreshTimeouts',
    ),
    destroyMainOverlayWindow: record('destroyMainOverlayWindow'),
    destroyModalOverlayWindow: record('destroyModalOverlayWindow'),
    destroyYomitanParserWindow: record('destroyYomitanParserWindow'),
    clearYomitanParserState: record('clearYomitanParserState'),
    stopWindowTracker: record('stopWindowTracker'),
    flushMpvLog: record('flushMpvLog'),
    destroyMpvSocket: record('destroyMpvSocket'),
    clearReconnectTimer: record('clearReconnectTimer'),
    destroySubtitleTimingTracker: record('destroySubtitleTimingTracker'),
    stopStatsServer: record('stopStatsServer'),
    destroyImmersionTracker: record('destroyImmersionTracker'),
    destroyAnkiIntegration: record('destroyAnkiIntegration'),
    destroyAnilistSetupWindow: record('destroyAnilistSetupWindow'),
    clearAnilistSetupWindow: record('clearAnilistSetupWindow'),
    destroyJellyfinSetupWindow: record('destroyJellyfinSetupWindow'),
    clearJellyfinSetupWindow: record('clearJellyfinSetupWindow'),
    destroyFirstRunSetupWindow: record('destroyFirstRunSetupWindow'),
    clearFirstRunSetupWindow: record('clearFirstRunSetupWindow'),
    destroyYomitanSettingsWindow: record('destroyYomitanSettingsWindow'),
    clearYomitanSettingsWindow: record('clearYomitanSettingsWindow'),
    stopJellyfinRemoteSession: record('stopJellyfinRemoteSession'),
    cleanupInternalSubtitleTrackCache: record('cleanupInternalSubtitleTrackCache'),
    cleanupYoutubeSubtitleTempDirs: record('cleanupYoutubeSubtitleTempDirs'),
    cleanupYoutubeMediaCache: record('cleanupYoutubeMediaCache'),
    cleanupRemoteMediaWindows: record('cleanupRemoteMediaWindows'),
    cleanupJellyfinSubtitleCache: record('cleanupJellyfinSubtitleCache'),
    stopDiscordPresenceService: record('stopDiscordPresenceService'),
    ...overrides,
  };
}

test('forced quit finalizes stats before exiting, even when finalization throws', async () => {
  for (const fails of [false, true]) {
    const calls: string[] = [];
    await createForceQuitHandler({
      destroyImmersionTracker: () => {
        calls.push('finalize');
        if (fails) throw new Error('flush failed');
      },
      logError: () => {
        calls.push('error');
      },
      exit: () => {
        calls.push('exit');
      },
    })();
    assert.deepEqual(calls, fails ? ['finalize', 'error', 'exit'] : ['finalize', 'exit']);
  }
});

test('on will quit cleanup handler flushes before teardown and awaits stats before dependents', async () => {
  const calls: string[] = [];
  const cleanup = createOnWillQuitCleanupHandler(
    makeCleanupDeps(calls, {
      stopStatsServer: async () => {
        await Promise.resolve();
        calls.push('stopStatsServer:complete');
      },
      destroyImmersionTracker: async () => {
        await Promise.resolve();
        calls.push('destroyImmersionTracker');
      },
    }),
  );

  await cleanup();
  assert.ok(calls.indexOf('flushMpvLog') < calls.indexOf('destroyMpvSocket'));
  assert.ok(calls.indexOf('stopStatsServer:complete') < calls.indexOf('destroyImmersionTracker'));
  assert.ok(calls.indexOf('destroyImmersionTracker') < calls.indexOf('destroyAnkiIntegration'));
});

test('on will quit cleanup handler cleans subtitle caches when stopping remote session fails', async () => {
  const calls: string[] = [];
  const cleanup = createOnWillQuitCleanupHandler(
    makeCleanupDeps(calls, {
      stopJellyfinRemoteSession: () => {
        calls.push('stopJellyfinRemoteSession');
        throw new Error('stop failed');
      },
    }),
  );

  await assert.rejects(cleanup(), /stop failed/);
  assert.deepEqual(calls.slice(calls.indexOf('stopJellyfinRemoteSession')), [
    'stopJellyfinRemoteSession',
    'cleanupJellyfinSubtitleCache',
    'cleanupInternalSubtitleTrackCache',
  ]);
});

test('forced quit exits when asynchronous stats finalization never settles', async () => {
  const calls: string[] = [];
  await createForceQuitHandler({
    destroyImmersionTracker: () => new Promise<void>(() => {}),
    logError: (error) => {
      assert.match(String(error), /Stats finalization timed out/);
      calls.push('timeout');
    },
    exit: () => calls.push('exit'),
  })();
  assert.deepEqual(calls, ['timeout', 'exit']);
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
