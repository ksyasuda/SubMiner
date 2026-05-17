import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createGetAnilistMediaGuessRuntimeStateHandler,
  createGetCurrentAnilistMediaKeyHandler,
  createRecordAnilistMediaDurationHandler,
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

test('reset anilist media tracking is idempotent', () => {
  const state = {
    mediaKey: 'old' as string | null,
    mediaDurationSec: 123 as number | null,
    mediaGuess: { title: 'guess' } as { title: string } | null,
    mediaGuessPromise: Promise.resolve(null) as Promise<unknown> | null,
    lastDurationProbeAtMs: 999,
  };

  const reset = createResetAnilistMediaTrackingHandler({
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

  reset('/new/media');
  const afterFirstReset = { ...state };
  reset('/new/media');

  assert.deepEqual(state, afterFirstReset);
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
  const state = {
    mediaKey: '/tmp/video.mkv' as string | null,
    mediaDurationSec: 240 as number | null,
    mediaGuess: { title: 'guess' } as { title: string } | null,
    mediaGuessPromise: Promise.resolve(null) as Promise<unknown> | null,
    lastDurationProbeAtMs: 321,
  };

  const resetGuessState = createResetAnilistMediaGuessStateHandler({
    setMediaGuess: (value) => {
      state.mediaGuess = value as { title: string } | null;
    },
    setMediaGuessPromise: (value) => {
      state.mediaGuessPromise = value;
    },
  });

  resetGuessState();
  assert.equal(state.mediaGuess, null);
  assert.equal(state.mediaGuessPromise, null);
  assert.equal(state.mediaKey, '/tmp/video.mkv');
  assert.equal(state.mediaDurationSec, 240);
  assert.equal(state.lastDurationProbeAtMs, 321);
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
