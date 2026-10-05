import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildAnilistAttemptKey,
  createMaybeRunAnilistPostWatchUpdateHandler,
  createProcessNextAnilistRetryUpdateHandler,
  rememberAnilistAttemptedUpdateKey,
} from './anilist-post-watch';

type PostWatchDeps = Parameters<typeof createMaybeRunAnilistPostWatchUpdateHandler>[0];
type RetryDeps = Parameters<typeof createProcessNextAnilistRetryUpdateHandler>[0];
type ProbeOptions = Parameters<PostWatchDeps['maybeProbeAnilistDuration']>[1];
type ProgressCall = {
  title: string;
  episode: number;
  season?: number | null;
  mediaId?: number | null;
};

const MEDIA_KEY = '/tmp/video.mkv';

function emptyOutcome() {
  return {
    inFlight: [] as boolean[],
    resets: [] as Array<string | null>,
    probes: [] as ProbeOptions[],
    guesses: 0,
    retryRuns: 0,
    tokenRefreshes: 0,
    updates: [] as ProgressCall[],
    enqueued: [] as Array<ProgressCall & { key: string }>,
    remembered: [] as string[],
    retrySuccesses: [] as string[],
    retryFailures: [] as Array<{ key: string; message: string }>,
    queueRefreshes: 0,
    osd: [] as string[],
    info: [] as string[],
    warn: [] as string[],
  };
}

// Defaults model a fully watched local episode 1 with a valid token. Default deps record
// into `outcome`; an overridden dep records nothing unless the test does it itself.
function makeDeps(overrides: Partial<PostWatchDeps> = {}) {
  const outcome = emptyOutcome();
  const attemptedKeys = new Set<string>();
  let inFlight = false;
  const deps: PostWatchDeps = {
    getInFlight: () => inFlight,
    setInFlight: (value) => {
      inFlight = value;
      outcome.inFlight.push(value);
    },
    getResolvedConfig: () => ({}),
    isAnilistTrackingEnabled: () => true,
    getCurrentMediaKey: () => MEDIA_KEY,
    hasMpvClient: () => true,
    getTrackedMediaKey: () => MEDIA_KEY,
    resetTrackedMedia: (mediaKey) => outcome.resets.push(mediaKey),
    getWatchedSeconds: () => 1000,
    maybeProbeAnilistDuration: async (_mediaKey, options) => {
      outcome.probes.push(options);
      return 1000;
    },
    ensureAnilistMediaGuess: async () => {
      outcome.guesses += 1;
      return { title: 'Show', season: null, episode: 1 };
    },
    hasAttemptedUpdateKey: (key) => attemptedKeys.has(key),
    processNextAnilistRetryUpdate: async () => {
      outcome.retryRuns += 1;
      return { ok: true, message: 'noop' };
    },
    refreshAnilistClientSecretState: async () => {
      outcome.tokenRefreshes += 1;
      return 'token';
    },
    enqueueRetry: (key, title, episode, season, mediaId) =>
      outcome.enqueued.push({ key, title, episode, season, mediaId }),
    markRetryFailure: (key, message) => outcome.retryFailures.push({ key, message }),
    markRetrySuccess: (key) => outcome.retrySuccesses.push(key),
    refreshRetryQueueState: () => {
      outcome.queueRefreshes += 1;
    },
    updateAnilistPostWatchProgress: async (_accessToken, title, episode, season, mediaId) => {
      outcome.updates.push({ title, episode, season, mediaId });
      return { status: 'updated', message: 'updated ok' };
    },
    rememberAttemptedUpdateKey: (key) => {
      attemptedKeys.add(key);
      outcome.remembered.push(key);
    },
    showMpvOsd: (message) => outcome.osd.push(message),
    logInfo: (message) => outcome.info.push(message),
    logWarn: (message) => outcome.warn.push(message),
    minWatchSeconds: 600,
    minWatchRatio: 0.85,
    ...overrides,
  };
  return { outcome, attemptedKeys, run: createMaybeRunAnilistPostWatchUpdateHandler(deps) };
}

test('buildAnilistAttemptKey formats media and episode', () => {
  assert.equal(buildAnilistAttemptKey('/tmp/video.mkv', 3), '/tmp/video.mkv::3');
  assert.equal(
    buildAnilistAttemptKey('https://example.com/Videos/item/stream?api_key=test-secret', 3),
    'jellyfin://example.com/item/item::3',
  );
});

test('rememberAnilistAttemptedUpdateKey evicts oldest beyond max size', () => {
  const set = new Set<string>(['a', 'b']);
  rememberAnilistAttemptedUpdateKey(set, 'c', 2);
  assert.deepEqual(Array.from(set), ['b', 'c']);
});

test('post-watch rejects empty media identities before attempted keys or update side effects', async () => {
  for (const mediaKey of [
    '  ',
    'stream?api_key=secret',
    'stream%3Fapi_key%3Dsecret',
    'https://[invalid',
  ]) {
    assert.equal(buildAnilistAttemptKey(mediaKey, 3), null);
    const { outcome, run } = makeDeps({
      getCurrentMediaKey: () => mediaKey,
      getTrackedMediaKey: () => mediaKey,
      hasAttemptedUpdateKey: () => assert.fail('invalid identity reached attempted-key lookup'),
    });
    await run();
    await run({ force: true });
    // Only gating ran: no retry, token, update, queue, or notification side effects.
    assert.deepEqual(outcome, {
      ...emptyOutcome(),
      inFlight: [true, false, true, false],
      probes: [{ force: false }],
      guesses: 2,
    });
  }
});

test('createProcessNextAnilistRetryUpdateHandler forwards the queued item and marks success', async () => {
  const updates: ProgressCall[] = [];
  const successes: string[] = [];
  const remembered: string[] = [];
  const errors: Array<string | null> = [];
  const deps: RetryDeps = {
    nextReady: () => ({ key: 'k1', title: 'Show', season: 2, mediaId: 108489, episode: 1 }),
    refreshRetryQueueState: () => {},
    setLastAttemptAt: () => {},
    setLastError: (value) => errors.push(value),
    refreshAnilistClientSecretState: async () => 'token',
    updateAnilistPostWatchProgress: async (_accessToken, title, episode, season, mediaId) => {
      updates.push({ title, episode, season, mediaId });
      return { status: 'updated', message: 'updated ok' };
    },
    markSuccess: (key) => successes.push(key),
    rememberAttemptedUpdateKey: (key) => remembered.push(key),
    markFailure: () => assert.fail('successful retry marked as failure'),
    logInfo: () => {},
    now: () => 1,
  };

  const result = await createProcessNextAnilistRetryUpdateHandler(deps)();
  assert.deepEqual(result, { ok: true, message: 'updated ok' });
  assert.deepEqual(updates, [{ title: 'Show', episode: 1, season: 2, mediaId: 108489 }]);
  assert.deepEqual(successes, ['k1']);
  assert.deepEqual(remembered, ['k1']);
  assert.deepEqual(errors, [null]);
});

for (const c of [
  { name: 'queues when token missing', pinnedMediaId: null },
  { name: 'queues the pinned media id for retry', pinnedMediaId: 108489 },
]) {
  test(`createMaybeRunAnilistPostWatchUpdateHandler ${c.name}`, async () => {
    const key = `${MEDIA_KEY}::1`;
    const { outcome, run } = makeDeps({
      refreshAnilistClientSecretState: async () => null,
      resolvePinnedAnilistMediaId: async () => c.pinnedMediaId,
    });

    await run();

    assert.deepEqual(outcome.updates, []);
    assert.deepEqual(outcome.enqueued, [
      { key, title: 'Show', episode: 1, season: null, mediaId: c.pinnedMediaId },
    ]);
    assert.deepEqual(outcome.retryFailures, [
      { key, message: 'cannot authenticate without anilist.accessToken' },
    ]);
    assert.deepEqual(outcome.osd, ['AniList: access token not configured']);
    assert.deepEqual(outcome.inFlight, [true, false]);
  });
}

test('createMaybeRunAnilistPostWatchUpdateHandler force-runs manual watched updates below threshold', async () => {
  const { outcome, run } = makeDeps({
    hasMpvClient: () => false,
    getWatchedSeconds: () => 0,
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: null, episode: 3 }),
  });

  await run({ force: true });

  assert.deepEqual(outcome.probes, []);
  assert.equal(outcome.updates.length, 1);
  assert.deepEqual(outcome.remembered, [`${MEDIA_KEY}::3`]);
  assert.deepEqual(outcome.osd, ['updated ok']);
});

test('createMaybeRunAnilistPostWatchUpdateHandler shows permanent AniList update errors without queueing retry', async () => {
  const message =
    'AniList update not possible: Show is not in your AniList Planning or Watching list.';
  let updateCalls = 0;
  const { outcome, run } = makeDeps({
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: null, episode: 2 }),
    updateAnilistPostWatchProgress: async () => {
      updateCalls += 1;
      return { status: 'error', retryable: false, message };
    },
  });

  await run();
  await run();

  assert.equal(updateCalls, 1);
  assert.deepEqual(outcome.enqueued, []);
  assert.deepEqual(outcome.retryFailures, []);
  assert.deepEqual(outcome.remembered, [`${MEDIA_KEY}::2`]);
  assert.equal(outcome.queueRefreshes, 1);
  assert.deepEqual(outcome.osd, [message]);
  assert.deepEqual(outcome.warn, [message]);
});

test('createMaybeRunAnilistPostWatchUpdateHandler uses provided watched seconds from time-position events', async () => {
  const { outcome, run } = makeDeps({
    getWatchedSeconds: () => 0,
    ensureAnilistMediaGuess: async () => ({ title: 'Show', season: 2, episode: 8 }),
  });

  await run({ watchedSeconds: 850 });

  assert.deepEqual(outcome.probes, [{ force: true }]);
  assert.deepEqual(outcome.updates, [{ title: 'Show', episode: 8, season: 2, mediaId: null }]);
  assert.deepEqual(outcome.remembered, [`${MEDIA_KEY}::8`]);
  assert.deepEqual(outcome.osd, ['updated ok']);
});

test('createMaybeRunAnilistPostWatchUpdateHandler blocks concurrent runs before async gating', async () => {
  let probes = 0;
  let resolveDuration!: (duration: number) => void;
  const durationPromise = new Promise<number>((resolve) => {
    resolveDuration = resolve;
  });
  const { outcome, run } = makeDeps({
    maybeProbeAnilistDuration: async () => {
      probes += 1;
      return await durationPromise;
    },
  });

  // The in-flight flag must be set synchronously, before the first await.
  const firstRun = run();
  assert.deepEqual(outcome.inFlight, [true]);
  assert.equal(probes, 1);

  await run();
  assert.deepEqual(outcome.inFlight, [true]);
  assert.equal(probes, 1);

  resolveDuration(1000);
  await firstRun;

  assert.equal(outcome.updates.length, 1);
  assert.deepEqual(outcome.inFlight, [true, false]);
});

test('createMaybeRunAnilistPostWatchUpdateHandler skips youtube playback entirely', async () => {
  const { outcome, run } = makeDeps({
    getCurrentMediaKey: () => 'https://www.youtube.com/watch?v=abc123',
  });

  await run();

  assert.deepEqual(outcome, emptyOutcome());
});

test('createMaybeRunAnilistPostWatchUpdateHandler notifies when retry already handled current attempt key', async () => {
  const { outcome, attemptedKeys, run } = makeDeps({
    processNextAnilistRetryUpdate: async () => {
      attemptedKeys.add(`${MEDIA_KEY}::1`);
      return { ok: true, message: 'retry ok' };
    },
  });

  await run();

  assert.deepEqual(outcome.osd, ['retry ok']);
  assert.equal(outcome.tokenRefreshes, 0);
  assert.deepEqual(outcome.updates, []);
  assert.deepEqual(outcome.enqueued, []);
  assert.deepEqual(outcome.retryFailures, []);
  assert.deepEqual(outcome.inFlight, [true, false]);
});

test('createMaybeRunAnilistPostWatchUpdateHandler passes the pinned override media id', async () => {
  const guess = { title: 'Show', season: 3, episode: 1 };
  const pinInputs: Array<Parameters<NonNullable<PostWatchDeps['resolvePinnedAnilistMediaId']>>[0]> =
    [];
  const { outcome, run } = makeDeps({
    ensureAnilistMediaGuess: async () => guess,
    getCurrentMediaTitle: () => 'Show S03E01.mkv',
    resolvePinnedAnilistMediaId: async (input) => {
      pinInputs.push(input);
      return 108489;
    },
  });

  await run();

  assert.deepEqual(outcome.updates, [{ title: 'Show', episode: 1, season: 3, mediaId: 108489 }]);
  // Resolved against the media this run captured, not whatever is playing now.
  assert.deepEqual(pinInputs, [{ mediaPath: MEDIA_KEY, mediaTitle: 'Show S03E01.mkv', guess }]);
});

test('createMaybeRunAnilistPostWatchUpdateHandler still updates when the override lookup throws', async () => {
  const { outcome, run } = makeDeps({
    resolvePinnedAnilistMediaId: async () => {
      throw new Error('store unreadable');
    },
  });

  await run();

  assert.deepEqual(
    outcome.updates.map((update) => update.mediaId),
    [null],
  );
  assert.equal(outcome.warn.length, 1);
  assert.match(outcome.warn[0]!, /override lookup failed/i);
});
