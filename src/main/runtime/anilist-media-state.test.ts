import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createGetAnilistMediaGuessRuntimeStateHandler,
  createGetCurrentAnilistMediaKeyHandler,
  createResetAnilistMediaGuessStateHandler,
  createResetAnilistMediaTrackingHandler,
  createSetAnilistMediaGuessRuntimeStateHandler,
} from './anilist-media-state';

test('get current anilist media key trims and normalizes empty path', () => {
  const getKey = createGetCurrentAnilistMediaKeyHandler({
    getCurrentMediaPath: () => '  /tmp/video.mkv  ',
  });
  const getEmptyKey = createGetCurrentAnilistMediaKeyHandler({
    getCurrentMediaPath: () => '   ',
  });

  assert.equal(getKey(), '/tmp/video.mkv');
  assert.equal(getEmptyKey(), null);
});

test('reset anilist media tracking clears duration/guess/probe state', () => {
  let mediaKey: string | null = 'old';
  let mediaDurationSec: number | null = 123;
  let mediaGuess: { title: string } | null = { title: 'guess' };
  let mediaGuessPromise: Promise<unknown> | null = Promise.resolve(null);
  let lastDurationProbeAtMs = 999;

  const reset = createResetAnilistMediaTrackingHandler({
    setMediaKey: (value) => {
      mediaKey = value;
    },
    setMediaDurationSec: (value) => {
      mediaDurationSec = value;
    },
    setMediaGuess: (value) => {
      mediaGuess = value as { title: string } | null;
    },
    setMediaGuessPromise: (value) => {
      mediaGuessPromise = value;
    },
    setLastDurationProbeAtMs: (value) => {
      lastDurationProbeAtMs = value;
    },
  });

  reset('/new/media');
  assert.equal(mediaKey, '/new/media');
  assert.equal(mediaDurationSec, null);
  assert.equal(mediaGuess, null);
  assert.equal(mediaGuessPromise, null);
  assert.equal(lastDurationProbeAtMs, 0);
});

test('get/set anilist media guess runtime state round-trips fields', () => {
  let state = {
    mediaKey: null as string | null,
    mediaDurationSec: null as number | null,
    mediaGuess: null as { title: string } | null,
    mediaGuessPromise: null as Promise<unknown> | null,
    lastDurationProbeAtMs: 0,
  };

  const setState = createSetAnilistMediaGuessRuntimeStateHandler({
    setMediaKey: (value) => {
      state.mediaKey = value;
    },
    setMediaDurationSec: (value) => {
      state.mediaDurationSec = value;
    },
    setMediaGuess: (value) => {
      state.mediaGuess = value as { title: string } | null;
    },
    setMediaGuessPromise: (value) => {
      state.mediaGuessPromise = value;
    },
    setLastDurationProbeAtMs: (value) => {
      state.lastDurationProbeAtMs = value;
    },
  });

  const getState = createGetAnilistMediaGuessRuntimeStateHandler({
    getMediaKey: () => state.mediaKey,
    getMediaDurationSec: () => state.mediaDurationSec,
    getMediaGuess: () => state.mediaGuess as never,
    getMediaGuessPromise: () => state.mediaGuessPromise as never,
    getLastDurationProbeAtMs: () => state.lastDurationProbeAtMs,
  });

  const nextPromise = Promise.resolve(null);
  setState({
    mediaKey: '/tmp/video.mkv',
    mediaDurationSec: 24,
    mediaGuess: { title: 'Title' } as never,
    mediaGuessPromise: nextPromise as never,
    lastDurationProbeAtMs: 321,
  });

  const roundTrip = getState();
  assert.equal(roundTrip.mediaKey, '/tmp/video.mkv');
  assert.equal(roundTrip.mediaDurationSec, 24);
  assert.deepEqual(roundTrip.mediaGuess, { title: 'Title' });
  assert.equal(roundTrip.mediaGuessPromise, nextPromise);
  assert.equal(roundTrip.lastDurationProbeAtMs, 321);
});

test('reset anilist media guess state clears guess and in-flight promise', () => {
  let mediaGuess: { title: string } | null = { title: 'guess' };
  let mediaGuessPromise: Promise<unknown> | null = Promise.resolve(null);

  const resetGuessState = createResetAnilistMediaGuessStateHandler({
    setMediaGuess: (value) => {
      mediaGuess = value as { title: string } | null;
    },
    setMediaGuessPromise: (value) => {
      mediaGuessPromise = value;
    },
  });

  resetGuessState();
  assert.equal(mediaGuess, null);
  assert.equal(mediaGuessPromise, null);
});
