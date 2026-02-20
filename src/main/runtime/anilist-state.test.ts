import test from 'node:test';
import assert from 'node:assert/strict';
import { createAnilistStateRuntime } from './anilist-state';
import type { AnilistRetryQueueState, AnilistSecretResolutionState } from '../state';

function createRuntime() {
  let clientState: AnilistSecretResolutionState = {
    status: 'resolved',
    source: 'stored',
    message: 'ok' as string | null,
    resolvedAt: 1000 as number | null,
    errorAt: null as number | null,
  };
  let queueState: AnilistRetryQueueState = {
    pending: 1,
    ready: 2,
    deadLetter: 3,
    lastAttemptAt: 2000 as number | null,
    lastError: 'none' as string | null,
  };
  let clearedStoredToken = false;
  let clearedCachedToken = false;

  const runtime = createAnilistStateRuntime({
    getClientSecretState: () => clientState,
    setClientSecretState: (next) => {
      clientState = next;
    },
    getRetryQueueState: () => queueState,
    setRetryQueueState: (next) => {
      queueState = next;
    },
    getUpdateQueueSnapshot: () => ({
      pending: 7,
      ready: 8,
      deadLetter: 9,
      lastAttemptAt: 3000,
      lastError: 'boom' as string | null,
    }),
    clearStoredToken: () => {
      clearedStoredToken = true;
    },
    clearCachedAccessToken: () => {
      clearedCachedToken = true;
    },
  });

  return {
    runtime,
    getClientState: () => clientState,
    getQueueState: () => queueState,
    getClearedStoredToken: () => clearedStoredToken,
    getClearedCachedToken: () => clearedCachedToken,
  };
}

test('setClientSecretState merges partial updates', () => {
  const harness = createRuntime();
  harness.runtime.setClientSecretState({
    status: 'error',
    source: 'none',
    errorAt: 4000,
  });

  assert.deepEqual(harness.getClientState(), {
    status: 'error',
    source: 'none',
    message: 'ok',
    resolvedAt: 1000,
    errorAt: 4000,
  });
});

test('refresh/get queue snapshot uses update queue snapshot', () => {
  const harness = createRuntime();
  const snapshot = harness.runtime.getQueueStatusSnapshot();

  assert.deepEqual(snapshot, {
    pending: 7,
    ready: 8,
    deadLetter: 9,
    lastAttemptAt: 3000,
    lastError: 'boom',
  });
  assert.deepEqual(harness.getQueueState(), snapshot);
});

test('clearTokenState resets token state and clears caches', () => {
  const harness = createRuntime();
  harness.runtime.clearTokenState();

  assert.equal(harness.getClearedStoredToken(), true);
  assert.equal(harness.getClearedCachedToken(), true);
  assert.deepEqual(harness.getClientState(), {
    status: 'not_checked',
    source: 'none',
    message: 'stored token cleared',
    resolvedAt: null,
    errorAt: null,
  });
});
