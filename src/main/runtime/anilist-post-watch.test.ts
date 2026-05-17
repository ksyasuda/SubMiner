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
    nextReady: () => ({ key: 'k1', title: 'Show', season: 2, episode: 1 }),
    refreshRetryQueueState: () => calls.push('refresh'),
    setLastAttemptAt: () => calls.push('attempt'),
    setLastError: (value) => calls.push(`error:${value ?? 'null'}`),
    refreshAnilistClientSecretState: async () => 'token',
    updateAnilistPostWatchProgress: async (_accessToken, _title, _episode, season) => ({
      status: 'updated',
      message: `updated ok:${season}`,
    }),
    markSuccess: () => calls.push('success'),
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    markFailure: () => calls.push('failure'),
    logInfo: () => calls.push('info'),
    now: () => 1,
  });

  const result = await handler();
  assert.deepEqual(result, { ok: true, message: 'updated ok:2' });
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
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: null, episode: 1 }),
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

test('createMaybeRunAnilistPostWatchUpdateHandler force-runs manual watched updates below threshold', async () => {
  const calls: string[] = [];
  const handler = createMaybeRunAnilistPostWatchUpdateHandler({
    getInFlight: () => false,
    setInFlight: (value) => calls.push(`inflight:${value}`),
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => '/tmp/video.mkv',
    hasMpvClient: () => false,
    getTrackedMediaKey: () => '/tmp/video.mkv',
    resetTrackedMedia: () => {},
    getWatchedSeconds: () => 0,
    maybeProbeAnilistDuration: async () => {
      calls.push('probe');
      return 1000;
    },
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: null, episode: 3 }),
    hasAttemptedUpdateKey: () => false,
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'noop' }),
    refreshAnilistClientSecretState: async () => 'token',
    enqueueRetry: () => calls.push('enqueue'),
    markRetryFailure: () => calls.push('mark-failure'),
    markRetrySuccess: () => calls.push('mark-success'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async () => {
      calls.push('update');
      return { status: 'updated', message: 'updated ok' };
    },
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
  });

  await handler({ force: true });

  assert.equal(calls.includes('probe'), false);
  assert.ok(calls.includes('update'));
  assert.ok(calls.includes('remember'));
  assert.ok(calls.includes('osd:updated ok'));
});

test('createMaybeRunAnilistPostWatchUpdateHandler shows permanent AniList update errors without queueing retry', async () => {
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
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: null, episode: 2 }),
    hasAttemptedUpdateKey: () => false,
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'noop' }),
    refreshAnilistClientSecretState: async () => 'token',
    enqueueRetry: () => calls.push('enqueue'),
    markRetryFailure: () => calls.push('mark-failure'),
    markRetrySuccess: () => calls.push('mark-success'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async () => ({
      status: 'error',
      retryable: false,
      message:
        'AniList update not possible: Show is not in your AniList Planning or Watching list.',
    }),
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
  });

  await handler();

  assert.equal(calls.includes('enqueue'), false);
  assert.equal(calls.includes('mark-failure'), false);
  assert.ok(calls.includes('refresh'));
  assert.ok(calls.some((call) => call.startsWith('osd:AniList update not possible')));
  assert.ok(calls.some((call) => call.startsWith('warn:AniList update not possible')));
});

test('createMaybeRunAnilistPostWatchUpdateHandler uses provided watched seconds from time-position events', async () => {
  const calls: string[] = [];
  let durationProbeOptions: unknown = null;
  const handler = createMaybeRunAnilistPostWatchUpdateHandler({
    getInFlight: () => false,
    setInFlight: (value) => calls.push(`inflight:${value}`),
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => '/tmp/video.mkv',
    hasMpvClient: () => true,
    getTrackedMediaKey: () => '/tmp/video.mkv',
    resetTrackedMedia: () => {},
    getWatchedSeconds: () => 0,
    maybeProbeAnilistDuration: async (_mediaKey, options) => {
      durationProbeOptions = options;
      return 1000;
    },
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: 2, episode: 8 }),
    hasAttemptedUpdateKey: () => false,
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'noop' }),
    refreshAnilistClientSecretState: async () => 'token',
    enqueueRetry: () => calls.push('enqueue'),
    markRetryFailure: () => calls.push('mark-failure'),
    markRetrySuccess: () => calls.push('mark-success'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async (_accessToken, _title, _episode, season) => {
      calls.push(`update:${season}`);
      return { status: 'updated', message: 'updated ok' };
    },
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
  });

  await handler({ watchedSeconds: 850 });

  assert.deepEqual(durationProbeOptions, { force: true });
  assert.ok(calls.includes('update:2'));
  assert.ok(calls.includes('remember'));
  assert.ok(calls.includes('osd:updated ok'));
});

test('createMaybeRunAnilistPostWatchUpdateHandler blocks concurrent runs before async gating', async () => {
  const calls: string[] = [];
  let inFlight = false;
  let resolveDuration!: (duration: number) => void;
  const durationPromise = new Promise<number>((resolve) => {
    resolveDuration = resolve;
  });
  const handler = createMaybeRunAnilistPostWatchUpdateHandler({
    getInFlight: () => inFlight,
    setInFlight: (value) => {
      inFlight = value;
      calls.push(`inflight:${value}`);
    },
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => '/tmp/video.mkv',
    hasMpvClient: () => true,
    getTrackedMediaKey: () => '/tmp/video.mkv',
    resetTrackedMedia: () => {},
    getWatchedSeconds: () => 1000,
    maybeProbeAnilistDuration: async () => {
      calls.push('probe');
      return await durationPromise;
    },
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: null, episode: 1 }),
    hasAttemptedUpdateKey: () => false,
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'noop' }),
    refreshAnilistClientSecretState: async () => 'token',
    enqueueRetry: () => calls.push('enqueue'),
    markRetryFailure: () => calls.push('mark-failure'),
    markRetrySuccess: () => calls.push('mark-success'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async () => {
      calls.push('update');
      return { status: 'updated', message: 'updated ok' };
    },
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
  });

  const firstRun = handler();
  assert.deepEqual(calls, ['inflight:true', 'probe']);

  await handler();
  assert.deepEqual(calls, ['inflight:true', 'probe']);

  resolveDuration(1000);
  await firstRun;

  assert.equal(calls.filter((call) => call === 'update').length, 1);
  assert.equal(calls.at(-1), 'inflight:false');
});

test('createMaybeRunAnilistPostWatchUpdateHandler skips youtube playback entirely', async () => {
  const calls: string[] = [];
  const handler = createMaybeRunAnilistPostWatchUpdateHandler({
    getInFlight: () => false,
    setInFlight: (value) => calls.push(`inflight:${value}`),
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => 'https://www.youtube.com/watch?v=abc123',
    hasMpvClient: () => true,
    getTrackedMediaKey: () => 'https://www.youtube.com/watch?v=abc123',
    resetTrackedMedia: () => calls.push('reset'),
    getWatchedSeconds: () => 1000,
    maybeProbeAnilistDuration: async () => {
      calls.push('probe');
      return 1000;
    },
    ensureAnilistMediaGuess: async () => {
      calls.push('guess');
      return { title: 'Show', season: null, episode: 1 };
    },
    hasAttemptedUpdateKey: () => false,
    processNextAnilistRetryUpdate: async () => {
      calls.push('process-retry');
      return { ok: true, message: 'noop' };
    },
    refreshAnilistClientSecretState: async () => {
      calls.push('refresh-token');
      return 'token';
    },
    enqueueRetry: () => calls.push('enqueue'),
    markRetryFailure: () => calls.push('mark-failure'),
    markRetrySuccess: () => calls.push('mark-success'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async () => {
      calls.push('update');
      return { status: 'updated', message: 'ok' };
    },
    rememberAttemptedUpdateKey: () => calls.push('remember'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
  });

  await handler();
  assert.deepEqual(calls, []);
});

test('createMaybeRunAnilistPostWatchUpdateHandler does not live-update after retry already handled current attempt key', async () => {
  const calls: string[] = [];
  const attemptedKeys = new Set<string>();
  const mediaKey = '/tmp/video.mkv';
  const attemptKey = buildAnilistAttemptKey(mediaKey, 1);
  const handler = createMaybeRunAnilistPostWatchUpdateHandler({
    getInFlight: () => false,
    setInFlight: (value) => calls.push(`inflight:${value}`),
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => mediaKey,
    hasMpvClient: () => true,
    getTrackedMediaKey: () => mediaKey,
    resetTrackedMedia: () => {},
    getWatchedSeconds: () => 1000,
    maybeProbeAnilistDuration: async () => 1000,
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: null, episode: 1 }),
    hasAttemptedUpdateKey: (key) => attemptedKeys.has(key),
    processNextAnilistRetryUpdate: async () => {
      attemptedKeys.add(attemptKey);
      calls.push('process-retry');
      return { ok: true, message: 'retry ok' };
    },
    refreshAnilistClientSecretState: async () => 'token',
    enqueueRetry: () => calls.push('enqueue'),
    markRetryFailure: () => calls.push('mark-failure'),
    markRetrySuccess: () => calls.push('mark-success'),
    refreshRetryQueueState: () => calls.push('refresh'),
    updateAnilistPostWatchProgress: async () => {
      calls.push('update');
      return { status: 'updated', message: 'updated ok' };
    },
    rememberAttemptedUpdateKey: (key) => {
      attemptedKeys.add(key);
      calls.push(`remember:${key}`);
    },
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
  });

  await handler();

  assert.equal(calls.includes('update'), false);
  assert.equal(calls.includes('enqueue'), false);
  assert.equal(calls.includes('mark-failure'), false);
  assert.deepEqual(calls, ['inflight:true', 'process-retry', 'inflight:false']);
});
