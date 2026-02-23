import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createEnsureAnilistMediaGuessHandler,
  createMaybeProbeAnilistDurationHandler,
  type AnilistMediaGuessRuntimeState,
} from './anilist-media-guess';

test('maybeProbeAnilistDuration updates state with probed duration', async () => {
  let state: AnilistMediaGuessRuntimeState = {
    mediaKey: '/tmp/video.mkv',
    mediaDurationSec: null,
    mediaGuess: null,
    mediaGuessPromise: null,
    lastDurationProbeAtMs: 0,
  };
  const probe = createMaybeProbeAnilistDurationHandler({
    getState: () => state,
    setState: (next) => {
      state = next;
    },
    durationRetryIntervalMs: 1000,
    now: () => 2000,
    requestMpvDuration: async () => 321,
    logWarn: () => {},
  });

  const duration = await probe('/tmp/video.mkv');
  assert.equal(duration, 321);
  assert.equal(state.mediaDurationSec, 321);
});

test('ensureAnilistMediaGuess memoizes in-flight guess promise', async () => {
  let state: AnilistMediaGuessRuntimeState = {
    mediaKey: '/tmp/video.mkv',
    mediaDurationSec: null,
    mediaGuess: null,
    mediaGuessPromise: null,
    lastDurationProbeAtMs: 0,
  };
  let calls = 0;
  const ensureGuess = createEnsureAnilistMediaGuessHandler({
    getState: () => state,
    setState: (next) => {
      state = next;
    },
    resolveMediaPathForJimaku: (value) => value,
    getCurrentMediaPath: () => '/tmp/video.mkv',
    getCurrentMediaTitle: () => 'Episode 1',
    guessAnilistMediaInfo: async () => {
      calls += 1;
      return { title: 'Show', episode: 1, source: 'guessit' };
    },
  });

  const [first, second] = await Promise.all([
    ensureGuess('/tmp/video.mkv'),
    ensureGuess('/tmp/video.mkv'),
  ]);
  assert.deepEqual(first, { title: 'Show', episode: 1, source: 'guessit' });
  assert.deepEqual(second, { title: 'Show', episode: 1, source: 'guessit' });
  assert.equal(calls, 1);
  assert.deepEqual(state.mediaGuess, { title: 'Show', episode: 1, source: 'guessit' });
  assert.equal(state.mediaGuessPromise, null);
});
