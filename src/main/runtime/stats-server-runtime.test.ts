import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createStatsServerRuntime,
  isSelfOwnedBackgroundStatsDaemonState,
  shouldClearAppStateStatsServerOnStop,
} from './stats-server-runtime';

test('detects self-owned background stats daemon state', () => {
  assert.equal(
    isSelfOwnedBackgroundStatsDaemonState({ pid: process.pid, port: 6969, startedAtMs: 1 }),
    true,
  );
});

test('stats server app-state reference should be cleared after private server stop', () => {
  assert.equal(shouldClearAppStateStatsServerOnStop({ hadStatsServer: true }), true);
});

test('stopBackgroundStatsServer clears stale state when daemon identity mismatches', async () => {
  const calls: string[] = [];
  const runtime = createStatsServerRuntime({
    userDataPath: '/tmp/subminer-stats-runtime-test',
    statsDistPath: '/tmp/stats-dist',
    getResolvedConfig: () => ({ stats: { serverPort: 5175 } }) as never,
    getImmersionTracker: () => null,
    setAppStateStatsServer: () => {},
    getMpvSocketPath: () => '/tmp/mpv.sock',
    getYomitanExt: () => null,
    getYomitanSession: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    getYomitanAnkiDeckName: async () => 'Mining',
    getAnilistRateLimiter: () => ({}) as never,
    resolveAnkiNoteId: (noteId) => noteId,
    trackDuplicateNoteIdsForNote: () => {},
    resolveSentenceSearchHeadwords: async () => [],
    ensureImmersionTrackerStarted: () => {},
    setStatsStartupInProgress: () => {},
    readBackgroundStatsServerState: () => ({ pid: 4242, port: 5175, startedAtMs: 1 }),
    removeBackgroundStatsServerState: () => {
      calls.push('removeBackgroundStatsServerState');
    },
    isBackgroundStatsServerProcessAlive: () => true,
    verifyBackgroundStatsServerIdentity: () => false,
    killProcess: () => {
      calls.push('killProcess');
    },
  });

  const result = await runtime.stopBackgroundStatsServer();

  assert.deepEqual(result, { ok: true, stale: true });
  assert.deepEqual(calls, ['removeBackgroundStatsServerState']);
});
