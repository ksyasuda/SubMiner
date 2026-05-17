import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler,
  createBuildProcessNextAnilistRetryUpdateMainDepsHandler,
} from './anilist-post-watch-main-deps';

test('process next anilist retry update main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildProcessNextAnilistRetryUpdateMainDepsHandler({
    nextReady: () => ({ key: 'k', title: 't', season: null, episode: 1 }),
    refreshRetryQueueState: () => calls.push('refresh'),
    setLastAttemptAt: () => calls.push('attempt'),
    setLastError: () => calls.push('error'),
    refreshAnilistClientSecretState: async () => 'token',
    updateAnilistPostWatchProgress: async (_accessToken, _title, _episode, season) => ({
      status: 'updated',
      message: `ok:${season}`,
    }),
    markSuccess: () => calls.push('success'),
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    markFailure: () => calls.push('failure'),
    logInfo: (message) => calls.push(`info:${message}`),
    now: () => 7,
  })();

  assert.deepEqual(deps.nextReady(), { key: 'k', title: 't', season: null, episode: 1 });
  deps.refreshRetryQueueState();
  deps.setLastAttemptAt(1);
  deps.setLastError('x');
  assert.equal(await deps.refreshAnilistClientSecretState(), 'token');
  assert.deepEqual(await deps.updateAnilistPostWatchProgress('token', 't', 1, 2), {
    status: 'updated',
    message: 'ok:2',
  });
  deps.markSuccess('k');
  deps.rememberAttemptedUpdateKey('k');
  deps.markFailure('k', 'bad');
  deps.logInfo('hello');
  assert.equal(deps.now(), 7);
  assert.deepEqual(calls, [
    'refresh',
    'attempt',
    'error',
    'success',
    'remember',
    'failure',
    'info:hello',
  ]);
});

test('maybe run anilist post watch update main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler({
    getInFlight: () => false,
    setInFlight: () => calls.push('in-flight'),
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => 'media',
    hasMpvClient: () => true,
    getTrackedMediaKey: () => 'media',
    resetTrackedMedia: () => calls.push('reset'),
    getWatchedSeconds: () => 100,
    maybeProbeAnilistDuration: async (_mediaKey, options) => {
      calls.push(`probe:${options?.force === true}`);
      return 120;
    },
    ensureAnilistMediaGuess: async () => ({ title: 'x', season: null, episode: 1 }),
    hasAttemptedUpdateKey: () => false,
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'ok' }),
    refreshAnilistClientSecretState: async () => 'token',
    enqueueRetry: (_key, _title, _episode, season) => calls.push(`enqueue:${season}`),
    markRetryFailure: () => calls.push('retry-fail'),
    markRetrySuccess: () => calls.push('retry-ok'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async (_accessToken, _title, _episode, season) => ({
      status: 'updated',
      message: `done:${season}`,
    }),
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    showMpvOsd: () => calls.push('osd'),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 5,
    minWatchRatio: 0.5,
  })();

  assert.equal(deps.getInFlight(), false);
  deps.setInFlight(true);
  assert.equal(deps.isAnilistTrackingEnabled(deps.getResolvedConfig()), true);
  assert.equal(deps.getCurrentMediaKey(), 'media');
  assert.equal(deps.hasMpvClient(), true);
  assert.equal(deps.getTrackedMediaKey(), 'media');
  deps.resetTrackedMedia('media');
  assert.equal(deps.getWatchedSeconds(), 100);
  assert.equal(await deps.maybeProbeAnilistDuration('media', { force: true }), 120);
  assert.deepEqual(await deps.ensureAnilistMediaGuess('media'), {
    title: 'x',
    season: null,
    episode: 1,
  });
  assert.equal(deps.hasAttemptedUpdateKey('k'), false);
  assert.deepEqual(await deps.processNextAnilistRetryUpdate(), { ok: true, message: 'ok' });
  assert.equal(await deps.refreshAnilistClientSecretState(), 'token');
  deps.enqueueRetry('k', 't', 1, 2);
  deps.markRetryFailure('k', 'bad');
  deps.markRetrySuccess('k');
  deps.refreshRetryQueueState();
  assert.deepEqual(await deps.updateAnilistPostWatchProgress('token', 't', 1, 2), {
    status: 'updated',
    message: 'done:2',
  });
  deps.rememberAttemptedUpdateKey('k');
  deps.showMpvOsd('ok');
  deps.logInfo('x');
  deps.logWarn('y');
  assert.equal(deps.minWatchSeconds, 5);
  assert.equal(deps.minWatchRatio, 0.5);
  assert.deepEqual(calls, [
    'in-flight',
    'reset',
    'probe:true',
    'enqueue:2',
    'retry-fail',
    'retry-ok',
    'refresh',
    'remember',
    'osd',
    'info:x',
    'warn:y',
  ]);
});
