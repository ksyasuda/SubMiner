import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createGetCurrentAnilistMediaKeyHandler,
  createRecordAnilistMediaDurationHandler,
  createResetAnilistMediaTrackingHandler,
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

test('get current anilist media key skips youtube playback urls', () => {
  const getYoutubeKey = createGetCurrentAnilistMediaKeyHandler({
    getCurrentMediaPath: () => ' https://www.youtube.com/watch?v=abc123 ',
  });
  const getShortYoutubeKey = createGetCurrentAnilistMediaKeyHandler({
    getCurrentMediaPath: () => 'https://youtu.be/abc123',
  });

  assert.equal(getYoutubeKey(), null);
  assert.equal(getShortYoutubeKey(), null);
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

test('record anilist media duration stores observed mpv duration for current media', () => {
  const existingPromise = Promise.resolve(null);
  let state = {
    mediaKey: '/tmp/video.mkv' as string | null,
    mediaDurationSec: null as number | null,
    mediaGuess: { title: 'guess' } as { title: string } | null,
    mediaGuessPromise: existingPromise as Promise<unknown> | null,
    lastDurationProbeAtMs: 321,
  };

  const recordDuration = createRecordAnilistMediaDurationHandler({
    getCurrentMediaKey: () => '/tmp/video.mkv',
    getState: () => state as never,
    setState: (nextState) => {
      state = nextState as never;
    },
  });

  recordDuration(1440);

  assert.equal(state.mediaDurationSec, 1440);
  assert.deepEqual(state.mediaGuess, { title: 'guess' });
  assert.equal(state.mediaGuessPromise, existingPromise);
  assert.equal(state.lastDurationProbeAtMs, 321);
});

test('record anilist media duration resets stale media state when media key changes', () => {
  let state = {
    mediaKey: '/tmp/old.mkv' as string | null,
    mediaDurationSec: 120 as number | null,
    mediaGuess: { title: 'old' } as { title: string } | null,
    mediaGuessPromise: Promise.resolve(null) as Promise<unknown> | null,
    lastDurationProbeAtMs: 321,
  };

  const recordDuration = createRecordAnilistMediaDurationHandler({
    getCurrentMediaKey: () => '/tmp/new.mkv',
    getState: () => state as never,
    setState: (nextState) => {
      state = nextState as never;
    },
  });

  recordDuration(1440);

  assert.deepEqual(state, {
    mediaKey: '/tmp/new.mkv',
    mediaDurationSec: 1440,
    mediaGuess: null,
    mediaGuessPromise: null,
    lastDurationProbeAtMs: 0,
  });
});
