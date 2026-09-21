import assert from 'node:assert/strict';
import test, { after } from 'node:test';
import { DEFAULT_CONFIG } from '../../config';
import { ImmersionTrackerService } from '../../core/services/immersion-tracker-service';
import { createAnilistRateLimiter } from '../../core/services/anilist/rate-limiter';
import {
  createStatsServerRuntime,
  isSelfOwnedBackgroundStatsDaemonState,
  type StatsServerRuntimeDeps,
} from './stats-server-runtime';
import type { StatsServer } from '../../core/services/stats-server';
import type { BackgroundStatsServerState } from './stats-daemon';

function createDeferred<T>() {
  let settle: ((value: T) => void) | null = null;
  let fail: ((error: unknown) => void) | null = null;
  const promise = new Promise<T>((resolve, reject) => {
    settle = resolve;
    fail = reject;
  });
  return {
    promise,
    resolve(value: T): void {
      if (!settle) throw new Error('deferred promise is unavailable');
      settle(value);
    },
    reject(error: unknown): void {
      if (!fail) throw new Error('deferred promise is unavailable');
      fail(error);
    },
  };
}

function createRuntimeHarness(
  startServer: NonNullable<StatsServerRuntimeDeps['startServer']>,
  backgroundState: BackgroundStatsServerState | null = null,
) {
  const appStateValues: Array<StatsServer | null> = [];
  const tracker = new ImmersionTrackerService({ dbPath: ':memory:' });
  after(() => tracker.destroy());
  const runtime = createStatsServerRuntime({
    userDataPath: '/tmp/subminer-stats-runtime-test',
    statsDistPath: '/tmp/stats-dist',
    getResolvedConfig: () => ({
      ...DEFAULT_CONFIG,
      stats: { ...DEFAULT_CONFIG.stats, serverPort: 5175 },
    }),
    getImmersionTracker: () => tracker,
    setAppStateStatsServer: (server) => {
      appStateValues.push(server);
    },
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
    getAnilistRateLimiter: () => createAnilistRateLimiter(),
    resolveAnkiNoteId: (noteId) => noteId,
    trackDuplicateNoteIdsForNote: () => {},
    resolveSentenceSearchHeadwords: async () => [],
    ensureImmersionTrackerStarted: () => {},
    setStatsStartupInProgress: () => {},
    readBackgroundStatsServerState: () => backgroundState,
    removeBackgroundStatsServerState: () => {},
    isBackgroundStatsServerProcessAlive: () => false,
    startServer,
  });
  return { runtime, appStateValues };
}

test('detects self-owned background stats daemon state', () => {
  assert.equal(
    isSelfOwnedBackgroundStatsDaemonState({ pid: process.pid, port: 6969, startedAtMs: 1 }),
    true,
  );
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

test('concurrent stats startup requests share one pending server', async () => {
  const deferred = createDeferred<StatsServer>();
  let startCalls = 0;
  const server: StatsServer = { close: async () => {} };
  const { runtime, appStateValues } = createRuntimeHarness(() => {
    startCalls += 1;
    return deferred.promise;
  });

  const first = runtime.ensureStatsServerStarted();
  const second = runtime.ensureStatsServerStarted();
  assert.equal(startCalls, 1);
  assert.deepEqual(appStateValues, []);

  deferred.resolve(server);
  assert.deepEqual(await Promise.all([first, second]), [
    { url: 'http://127.0.0.1:5175', source: 'local' },
    { url: 'http://127.0.0.1:5175', source: 'local' },
  ]);
  assert.deepEqual(appStateValues, [server]);
});

test('failed stats startup remains recoverable on the next request', async () => {
  const first = createDeferred<StatsServer>();
  const second = createDeferred<StatsServer>();
  const attempts = [first, second];
  let startCalls = 0;
  const server: StatsServer = { close: async () => {} };
  const { runtime, appStateValues } = createRuntimeHarness(() => {
    const attempt = attempts[startCalls];
    startCalls += 1;
    if (!attempt) throw new Error('unexpected startup attempt');
    return attempt.promise;
  });

  const failedStartup = runtime.ensureStatsServerStarted();
  first.reject(Object.assign(new Error('address in use'), { code: 'EADDRINUSE' }));
  await assert.rejects(failedStartup, /address in use/);

  const retry = runtime.ensureStatsServerStarted();
  second.resolve(server);
  assert.deepEqual(await retry, { url: 'http://127.0.0.1:5175', source: 'local' });
  assert.equal(startCalls, 2);
  assert.deepEqual(appStateValues, [null, server]);
});

test('shutdown cancels pending startup and closes the late server', async () => {
  const deferred = createDeferred<StatsServer>();
  let closeCalls = 0;
  const server: StatsServer = {
    close: async () => {
      closeCalls += 1;
    },
  };
  const { runtime, appStateValues } = createRuntimeHarness(() => deferred.promise);

  const startup = runtime.ensureStatsServerStarted();
  const shutdown = runtime.stopStatsServer();
  deferred.resolve(server);

  await assert.rejects(startup, /startup was cancelled/);
  await shutdown;
  assert.equal(closeCalls, 1);
  assert.deepEqual(appStateValues, [null, null]);
});

test('stopping a self-owned background server closes its local handle', async () => {
  let closeCalls = 0;
  const server: StatsServer = {
    close: async () => {
      closeCalls += 1;
    },
  };
  const { runtime } = createRuntimeHarness(async () => server, {
    pid: process.pid,
    port: 5175,
    startedAtMs: 1,
  });
  await runtime.ensureStatsServerStarted();

  assert.deepEqual(await runtime.stopBackgroundStatsServer(), { ok: true, stale: false });
  assert.equal(closeCalls, 1);
});

test('background stop leaves a foreground-only server available', async () => {
  let closeCalls = 0;
  let startCalls = 0;
  const { runtime } = createRuntimeHarness(async () => {
    startCalls += 1;
    return {
      close: async () => {
        closeCalls += 1;
      },
    };
  });
  const foreground = await runtime.ensureStatsServerStarted();
  assert.deepEqual(await runtime.stopBackgroundStatsServer(), { ok: true, stale: true });
  assert.equal(closeCalls, 0);
  assert.deepEqual(await runtime.ensureStatsServerStarted(), foreground);
  assert.equal(startCalls, 1);
  await runtime.stopStatsServer();
});

test('background stop leaves a pending foreground-only startup alone', async () => {
  const deferred = createDeferred<StatsServer>();
  const { runtime } = createRuntimeHarness(() => deferred.promise);
  const startup = runtime.ensureStatsServerStarted();
  assert.deepEqual(await runtime.stopBackgroundStatsServer(), { ok: true, stale: true });
  deferred.resolve({ close: async () => {} });
  assert.deepEqual(await startup, { url: 'http://127.0.0.1:5175', source: 'local' });
  await runtime.stopStatsServer();
});

test('a startup requested during shutdown waits and then restarts', async () => {
  const closeDeferred = createDeferred<void>();
  const firstServer: StatsServer = { close: () => closeDeferred.promise };
  const secondServer: StatsServer = { close: async () => {} };
  const servers = [firstServer, secondServer];
  let startCalls = 0;
  const { runtime } = createRuntimeHarness(async () => {
    const server = servers[startCalls];
    startCalls += 1;
    if (!server) throw new Error('unexpected startup attempt');
    return server;
  });
  await runtime.ensureStatsServerStarted();

  const shutdown = runtime.stopStatsServer();
  const restart = runtime.ensureStatsServerStarted();
  assert.equal(startCalls, 1);

  closeDeferred.resolve();
  await shutdown;
  assert.deepEqual(await restart, { url: 'http://127.0.0.1:5175', source: 'local' });
  assert.equal(startCalls, 2);
});

test('background stop cancels startup before daemon ownership is published', async () => {
  const deferred = createDeferred<StatsServer>();
  let closeCalls = 0;
  const { runtime, appStateValues } = createRuntimeHarness(() => deferred.promise);
  const startup = runtime.ensureBackgroundStatsServerStarted();
  const shutdown = runtime.stopBackgroundStatsServer();
  deferred.resolve({
    close: async () => {
      closeCalls += 1;
    },
  });
  await assert.rejects(startup, /startup was cancelled/);
  assert.deepEqual(await shutdown, { ok: true, stale: false });
  assert.equal(closeCalls, 1);
  assert.equal(appStateValues.at(-1), null);
});
