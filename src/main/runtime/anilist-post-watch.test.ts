import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildAnilistAttemptKey,
  createMaybeRunAnilistPostWatchUpdateHandler,
  createProcessNextAnilistRetryUpdateHandler,
  rememberAnilistAttemptedUpdateKey,
} from './anilist-post-watch';

test('buildAnilistAttemptKey formats media and episode', () => {
  assert.equal(buildAnilistAttemptKey('/tmp/video.mkv', 3), '/tmp/video.mkv::3');
});

test('rememberAnilistAttemptedUpdateKey evicts oldest beyond max size', () => {
  const set = new Set<string>(['a', 'b']);
  rememberAnilistAttemptedUpdateKey(set, 'c', 2);
  assert.deepEqual(Array.from(set), ['b', 'c']);
});

test('createProcessNextAnilistRetryUpdateHandler handles successful retry', async () => {
  const calls: string[] = [];
  const handler = createProcessNextAnilistRetryUpdateHandler({
    nextReady: () => ({ key: 'k1', title: 'Show', episode: 1 }),
    refreshRetryQueueState: () => calls.push('refresh'),
    setLastAttemptAt: () => calls.push('attempt'),
    setLastError: (value) => calls.push(`error:${value ?? 'null'}`),
    refreshAnilistClientSecretState: async () => 'token',
    updateAnilistPostWatchProgress: async () => ({ status: 'updated', message: 'updated ok' }),
    markSuccess: () => calls.push('success'),
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    markFailure: () => calls.push('failure'),
    logInfo: () => calls.push('info'),
    now: () => 1,
  });

  const result = await handler();
  assert.deepEqual(result, { ok: true, message: 'updated ok' });
  assert.ok(calls.includes('success'));
  assert.ok(calls.includes('remember'));
});

test('createMaybeRunAnilistPostWatchUpdateHandler queues when token missing', async () => {
  const calls: string[] = [];
  const handler = createMaybeRunAnilistPostWatchUpdateHandler({
    getInFlight: () => false,
    setInFlight: (value) => calls.push(`inflight:${value}`),
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => '/tmp/video.mkv',
    hasMpvClient: () => true,
    getTrackedMediaKey: () => '/tmp/video.mkv',
    resetTrackedMedia: () => {},
    getWatchedSeconds: () => 1000,
    maybeProbeAnilistDuration: async () => 1000,
    ensureAnilistMediaGuess: async () => ({ title: 'Show', episode: 1 }),
    hasAttemptedUpdateKey: () => false,
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'noop' }),
    refreshAnilistClientSecretState: async () => null,
    enqueueRetry: () => calls.push('enqueue'),
    markRetryFailure: () => calls.push('mark-failure'),
    markRetrySuccess: () => calls.push('mark-success'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async () => ({ status: 'updated', message: 'ok' }),
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
  });

  await handler();
  assert.ok(calls.includes('enqueue'));
  assert.ok(calls.includes('mark-failure'));
  assert.ok(calls.includes('osd:AniList: access token not configured'));
  assert.ok(calls.includes('inflight:true'));
  assert.ok(calls.includes('inflight:false'));
});
