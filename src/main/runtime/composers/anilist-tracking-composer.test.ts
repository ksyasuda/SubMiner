import assert from 'node:assert/strict';
import test from 'node:test';
import type { AnilistMediaGuess } from '../../../core/services/anilist/anilist-updater';
import { composeAnilistTrackingHandlers } from './anilist-tracking-composer';

test('composeAnilistTrackingHandlers returns callable handlers and forwards calls to deps', async () => {
  const refreshSavedTokens: string[] = [];
  let refreshCachedToken: string | null = null;

  let mediaKeyState: string | null = 'media-key';
  let mediaDurationSecState: number | null = null;
  let mediaGuessState: AnilistMediaGuess | null = null;
  let mediaGuessPromiseState: Promise<AnilistMediaGuess | null> | null = null;
  let lastDurationProbeAtMsState = 0;
  let requestMpvDurationCalls = 0;
  let guessAnilistMediaInfoCalls = 0;

  let retryUpdateCalls = 0;
  let maybeRunUpdateCalls = 0;

  const composed = composeAnilistTrackingHandlers({
    refreshClientSecretMainDeps: {
      getResolvedConfig: () => ({ anilist: { accessToken: 'refresh-token' } }),
      isAnilistTrackingEnabled: () => true,
      getCachedAccessToken: () => refreshCachedToken,
      setCachedAccessToken: (token) => {
        refreshCachedToken = token;
      },
      saveStoredToken: (token) => {
        refreshSavedTokens.push(token);
      },
      loadStoredToken: () => null,
      setClientSecretState: () => {},
      getAnilistSetupPageOpened: () => false,
      setAnilistSetupPageOpened: () => {},
      openAnilistSetupWindow: () => {},
      now: () => 100,
    },
    getCurrentMediaKeyMainDeps: {
      getCurrentMediaPath: () => '  media-key  ',
    },
    resetMediaTrackingMainDeps: {
      setMediaKey: (value) => {
        mediaKeyState = value;
      },
      setMediaDurationSec: (value) => {
        mediaDurationSecState = value;
      },
      setMediaGuess: (value) => {
        mediaGuessState = value;
      },
      setMediaGuessPromise: (value) => {
        mediaGuessPromiseState = value;
      },
      setLastDurationProbeAtMs: (value) => {
        lastDurationProbeAtMsState = value;
      },
    },
    getMediaGuessRuntimeStateMainDeps: {
      getMediaKey: () => mediaKeyState,
      getMediaDurationSec: () => mediaDurationSecState,
      getMediaGuess: () => mediaGuessState,
      getMediaGuessPromise: () => mediaGuessPromiseState,
      getLastDurationProbeAtMs: () => lastDurationProbeAtMsState,
    },
    setMediaGuessRuntimeStateMainDeps: {
      setMediaKey: (value) => {
        mediaKeyState = value;
      },
      setMediaDurationSec: (value) => {
        mediaDurationSecState = value;
      },
      setMediaGuess: (value) => {
        mediaGuessState = value;
      },
      setMediaGuessPromise: (value) => {
        mediaGuessPromiseState = value;
      },
      setLastDurationProbeAtMs: (value) => {
        lastDurationProbeAtMsState = value;
      },
    },
    resetMediaGuessStateMainDeps: {
      setMediaGuess: (value) => {
        mediaGuessState = value;
      },
      setMediaGuessPromise: (value) => {
        mediaGuessPromiseState = value;
      },
    },
    maybeProbeDurationMainDeps: {
      getState: () => ({
        mediaKey: mediaKeyState,
        mediaDurationSec: mediaDurationSecState,
        mediaGuess: mediaGuessState,
        mediaGuessPromise: mediaGuessPromiseState,
        lastDurationProbeAtMs: lastDurationProbeAtMsState,
      }),
      setState: (state) => {
        mediaKeyState = state.mediaKey;
        mediaDurationSecState = state.mediaDurationSec;
        mediaGuessState = state.mediaGuess;
        mediaGuessPromiseState = state.mediaGuessPromise;
        lastDurationProbeAtMsState = state.lastDurationProbeAtMs;
      },
      durationRetryIntervalMs: 0,
      now: () => 1000,
      requestMpvDuration: async () => {
        requestMpvDurationCalls += 1;
        return 120;
      },
      logWarn: () => {},
    },
    ensureMediaGuessMainDeps: {
      getState: () => ({
        mediaKey: mediaKeyState,
        mediaDurationSec: mediaDurationSecState,
        mediaGuess: mediaGuessState,
        mediaGuessPromise: mediaGuessPromiseState,
        lastDurationProbeAtMs: lastDurationProbeAtMsState,
      }),
      setState: (state) => {
        mediaKeyState = state.mediaKey;
        mediaDurationSecState = state.mediaDurationSec;
        mediaGuessState = state.mediaGuess;
        mediaGuessPromiseState = state.mediaGuessPromise;
        lastDurationProbeAtMsState = state.lastDurationProbeAtMs;
      },
      resolveMediaPathForJimaku: (value) => value,
      getCurrentMediaPath: () => '/tmp/media.mkv',
      getCurrentMediaTitle: () => 'Episode title',
      guessAnilistMediaInfo: async () => {
        guessAnilistMediaInfoCalls += 1;
        return { title: 'Episode title', season: null, episode: 7, source: 'guessit' };
      },
    },
    processNextRetryUpdateMainDeps: {
      nextReady: () => ({ key: 'retry-key', title: 'Retry title', season: null, episode: 1 }),
      refreshRetryQueueState: () => {},
      setLastAttemptAt: () => {},
      setLastError: () => {},
      refreshAnilistClientSecretState: async () => 'retry-token',
      updateAnilistPostWatchProgress: async () => {
        retryUpdateCalls += 1;
        return { status: 'updated', message: 'ok' };
      },
      markSuccess: () => {},
      rememberAttemptedUpdateKey: () => {},
      markFailure: () => {},
      logInfo: () => {},
      now: () => 1,
    },
    maybeRunPostWatchUpdateMainDeps: {
      getInFlight: () => false,
      setInFlight: () => {},
      getResolvedConfig: () => ({ tracking: true }),
      isAnilistTrackingEnabled: () => true,
      getCurrentMediaKey: () => 'media-key',
      hasMpvClient: () => true,
      getTrackedMediaKey: () => 'media-key',
      resetTrackedMedia: () => {},
      getWatchedSeconds: () => 500,
      maybeProbeAnilistDuration: async () => 600,
      ensureAnilistMediaGuess: async () => ({
        title: 'Episode title',
        season: null,
        episode: 2,
        source: 'guessit',
      }),
      hasAttemptedUpdateKey: () => false,
      processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'ok' }),
      refreshAnilistClientSecretState: async () => 'run-token',
      enqueueRetry: () => {},
      markRetryFailure: () => {},
      markRetrySuccess: () => {},
      refreshRetryQueueState: () => {},
      updateAnilistPostWatchProgress: async () => {
        maybeRunUpdateCalls += 1;
        return { status: 'updated', message: 'updated from maybeRun' };
      },
      rememberAttemptedUpdateKey: () => {},
      showMpvOsd: () => {},
      logInfo: () => {},
      logWarn: () => {},
      minWatchSeconds: 10,
      minWatchRatio: 0.5,
    },
  });

  assert.equal(typeof composed.refreshAnilistClientSecretState, 'function');
  assert.equal(typeof composed.getCurrentAnilistMediaKey, 'function');
  assert.equal(typeof composed.resetAnilistMediaTracking, 'function');
  assert.equal(typeof composed.getAnilistMediaGuessRuntimeState, 'function');
  assert.equal(typeof composed.setAnilistMediaGuessRuntimeState, 'function');
  assert.equal(typeof composed.resetAnilistMediaGuessState, 'function');
  assert.equal(typeof composed.maybeProbeAnilistDuration, 'function');
  assert.equal(typeof composed.ensureAnilistMediaGuess, 'function');
  assert.equal(typeof composed.processNextAnilistRetryUpdate, 'function');
  assert.equal(typeof composed.maybeRunAnilistPostWatchUpdate, 'function');

  const refreshed = await composed.refreshAnilistClientSecretState({ force: true });
  assert.equal(refreshed, 'refresh-token');
  assert.deepEqual(refreshSavedTokens, ['refresh-token']);

  assert.equal(composed.getCurrentAnilistMediaKey(), 'media-key');
  composed.resetAnilistMediaTracking('next-key');
  assert.equal(mediaKeyState, 'next-key');
  assert.equal(mediaDurationSecState, null);

  composed.setAnilistMediaGuessRuntimeState({
    mediaKey: 'media-key',
    mediaDurationSec: 90,
    mediaGuess: { title: 'Known', season: null, episode: 3, source: 'fallback' },
    mediaGuessPromise: null,
    lastDurationProbeAtMs: 11,
  });
  assert.equal(composed.getAnilistMediaGuessRuntimeState().mediaDurationSec, 90);

  composed.resetAnilistMediaGuessState();
  assert.equal(composed.getAnilistMediaGuessRuntimeState().mediaGuess, null);

  mediaKeyState = 'media-key';
  mediaDurationSecState = null;
  const probedDuration = await composed.maybeProbeAnilistDuration('media-key');
  assert.equal(probedDuration, 120);
  assert.equal(requestMpvDurationCalls, 1);

  mediaGuessState = null;
  await composed.ensureAnilistMediaGuess('media-key');
  assert.equal(guessAnilistMediaInfoCalls, 1);

  const retryResult = await composed.processNextAnilistRetryUpdate();
  assert.deepEqual(retryResult, { ok: true, message: 'ok' });
  assert.equal(retryUpdateCalls, 1);

  await composed.maybeRunAnilistPostWatchUpdate();
  assert.equal(maybeRunUpdateCalls, 1);
});
