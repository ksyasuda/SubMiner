import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildEnsureAnilistMediaGuessMainDepsHandler,
  createBuildMaybeProbeAnilistDurationMainDepsHandler,
} from './anilist-media-guess-main-deps';

test('maybe probe anilist duration main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildMaybeProbeAnilistDurationMainDepsHandler({
    getState: () => ({
      mediaKey: 'm',
      mediaDurationSec: null,
      mediaGuess: null,
      mediaGuessPromise: null,
      lastDurationProbeAtMs: 0,
    }),
    setState: () => calls.push('set-state'),
    durationRetryIntervalMs: 1000,
    now: () => 42,
    requestMpvDuration: async () => 3600,
    logWarn: (message) => calls.push(`warn:${message}`),
  })();

  assert.equal(deps.durationRetryIntervalMs, 1000);
  assert.equal(deps.now(), 42);
  assert.equal(await deps.requestMpvDuration(), 3600);
  deps.setState({
    mediaKey: 'm',
    mediaDurationSec: 100,
    mediaGuess: null,
    mediaGuessPromise: null,
    lastDurationProbeAtMs: 0,
  });
  deps.logWarn('oops', null);
  assert.deepEqual(calls, ['set-state', 'warn:oops']);
});

test('ensure anilist media guess main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildEnsureAnilistMediaGuessMainDepsHandler({
    getState: () => ({
      mediaKey: 'm',
      mediaDurationSec: null,
      mediaGuess: null,
      mediaGuessPromise: null,
      lastDurationProbeAtMs: 0,
    }),
    setState: () => calls.push('set-state'),
    resolveMediaPathForJimaku: (path) => {
      calls.push('resolve');
      return path;
    },
    getCurrentMediaPath: () => '/tmp/video.mkv',
    getCurrentMediaTitle: () => 'title',
    guessAnilistMediaInfo: async () => {
      calls.push('guess');
      return { title: 'title', season: null, episode: 1, source: 'fallback' };
    },
  })();

  assert.equal(deps.getCurrentMediaPath(), '/tmp/video.mkv');
  assert.equal(deps.getCurrentMediaTitle(), 'title');
  assert.equal(deps.resolveMediaPathForJimaku('/tmp/video.mkv'), '/tmp/video.mkv');
  assert.deepEqual(await deps.guessAnilistMediaInfo('/tmp/video.mkv', 'title'), {
    title: 'title',
    season: null,
    episode: 1,
    source: 'fallback',
  });
  deps.setState({
    mediaKey: 'm',
    mediaDurationSec: null,
    mediaGuess: null,
    mediaGuessPromise: null,
    lastDurationProbeAtMs: 0,
  });
  assert.deepEqual(calls, ['resolve', 'guess', 'set-state']);
});
