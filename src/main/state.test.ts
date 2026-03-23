import test from 'node:test';
import assert from 'node:assert/strict';

import { parseArgs } from '../cli/args';
import {
  applyStartupState,
  createAppState,
  createInitialAnilistMediaGuessRuntimeState,
  createInitialAnilistUpdateInFlightState,
  transitionAnilistClientSecretState,
  transitionAnilistMediaGuessRuntimeState,
  transitionAnilistRetryQueueLastAttemptAt,
  transitionAnilistRetryQueueLastError,
  transitionAnilistUpdateInFlightState,
} from './state';

test('transitionAnilistClientSecretState replaces state object', () => {
  const current = {
    status: 'not_checked',
    source: 'none',
    message: null,
    resolvedAt: null,
    errorAt: null,
  } as const;
  const next = {
    status: 'resolved',
    source: 'stored',
    message: 'ok',
    resolvedAt: 123,
    errorAt: null,
  } as const;

  const transitioned = transitionAnilistClientSecretState(current, next);

  assert.deepEqual(transitioned, next);
  assert.equal(transitioned, next);
});

test('retry queue metadata transitions preserve queue counts', () => {
  const queue = {
    pending: 2,
    ready: 1,
    deadLetter: 4,
    lastAttemptAt: null,
    lastError: null,
  };

  const attempted = transitionAnilistRetryQueueLastAttemptAt(queue, 999);
  const failed = transitionAnilistRetryQueueLastError(attempted, 'boom');

  assert.deepEqual(attempted, {
    pending: 2,
    ready: 1,
    deadLetter: 4,
    lastAttemptAt: 999,
    lastError: null,
  });
  assert.deepEqual(failed, {
    pending: 2,
    ready: 1,
    deadLetter: 4,
    lastAttemptAt: 999,
    lastError: 'boom',
  });
  assert.notEqual(attempted, queue);
  assert.notEqual(failed, attempted);
});

test('transitionAnilistMediaGuessRuntimeState applies partial updates', () => {
  const current = createInitialAnilistMediaGuessRuntimeState();
  const promise = Promise.resolve(null);

  const transitioned = transitionAnilistMediaGuessRuntimeState(current, {
    mediaKey: '/tmp/media.mkv',
    mediaGuessPromise: promise,
    lastDurationProbeAtMs: 500,
  });

  assert.deepEqual(transitioned, {
    mediaKey: '/tmp/media.mkv',
    mediaDurationSec: null,
    mediaGuess: null,
    mediaGuessPromise: promise,
    lastDurationProbeAtMs: 500,
  });
  assert.notEqual(transitioned, current);
});

test('transitionAnilistUpdateInFlightState updates inFlight only', () => {
  const current = createInitialAnilistUpdateInFlightState();
  const transitioned = transitionAnilistUpdateInFlightState(current, true);

  assert.deepEqual(current, { inFlight: false });
  assert.deepEqual(transitioned, { inFlight: true });
  assert.notEqual(transitioned, current);
});

test('applyStartupState does not mark youtube playback flow pending from startup args alone', () => {
  const appState = createAppState({
    mpvSocketPath: '/tmp/mpv.sock',
    texthookerPort: 4000,
  });

  applyStartupState(appState, {
    initialArgs: parseArgs(['--youtube-play', 'https://www.youtube.com/watch?v=video123']),
    mpvSocketPath: '/tmp/mpv.sock',
    texthookerPort: 4000,
    backendOverride: null,
    autoStartOverlay: false,
    texthookerOnlyMode: false,
    backgroundMode: false,
  });

  assert.equal(appState.youtubePlaybackFlowPending, false);
});
