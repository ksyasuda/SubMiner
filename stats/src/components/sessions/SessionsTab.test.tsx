import assert from 'node:assert/strict';
import test from 'node:test';
import type { SessionBucket } from '../../lib/session-grouping';
import type { SessionSummary } from '../../types/stats';
import { buildBucketDeleteHandler } from './SessionsTab';

function makeSession(over: Partial<SessionSummary>): SessionSummary {
  return {
    sessionId: 1,
    videoId: 100,
    canonicalTitle: 'Episode 1',
    startedAtMs: 1_000_000,
    endedAtMs: 1_060_000,
    activeWatchedMs: 60_000,
    cardsMined: 1,
    linesSeen: 10,
    lookupCount: 5,
    lookupHits: 3,
    knownWordsSeen: 5,
    ...over,
  } as SessionSummary;
}

function makeBucket(sessions: SessionSummary[]): SessionBucket {
  const sorted = [...sessions].sort((a, b) => b.startedAtMs - a.startedAtMs);
  return {
    key: `v-${sorted[0]!.videoId}`,
    videoId: sorted[0]!.videoId ?? null,
    sessions: sorted,
    totalActiveMs: sorted.reduce((s, x) => s + x.activeWatchedMs, 0),
    totalCardsMined: sorted.reduce((s, x) => s + x.cardsMined, 0),
    representativeSession: sorted[0]!,
  };
}

test('buildBucketDeleteHandler deletes every session in the bucket when confirm returns true', async () => {
  let deleted: number[] | null = null;
  let onSuccessCalledWith: number[] | null = null;
  let onErrorCalled = false;

  const bucket = makeBucket([
    makeSession({ sessionId: 11, startedAtMs: 2_000_000 }),
    makeSession({ sessionId: 22, startedAtMs: 3_000_000 }),
    makeSession({ sessionId: 33, startedAtMs: 4_000_000 }),
  ]);

  const handler = buildBucketDeleteHandler({
    bucket,
    apiClient: {
      deleteSessions: async (ids: number[]) => {
        deleted = ids;
      },
    },
    confirm: (title, count) => {
      assert.equal(title, 'Episode 1');
      assert.equal(count, 3);
      return true;
    },
    onSuccess: (ids) => {
      onSuccessCalledWith = ids;
    },
    onError: () => {
      onErrorCalled = true;
    },
  });

  await handler();

  assert.deepEqual(deleted, [33, 22, 11]);
  assert.deepEqual(onSuccessCalledWith, [33, 22, 11]);
  assert.equal(onErrorCalled, false);
});

test('buildBucketDeleteHandler signals deleted session IDs after confirm, before deleting', async () => {
  const events: string[] = [];
  let startedIds: number[] | null = null;

  const bucket = makeBucket([
    makeSession({ sessionId: 11 }),
    makeSession({ sessionId: 22 }),
    makeSession({ sessionId: 33 }),
  ]);

  const handler = buildBucketDeleteHandler({
    bucket,
    apiClient: {
      deleteSessions: async () => {
        events.push('delete');
      },
    },
    confirm: () => {
      events.push('confirm');
      return true;
    },
    onStart: (ids) => {
      startedIds = ids;
      events.push('start');
    },
    onSuccess: () => {
      events.push('success');
    },
    onError: () => {
      events.push('error');
    },
  });

  await handler();

  assert.deepEqual(events, ['confirm', 'start', 'delete', 'success']);
  assert.deepEqual(startedIds, [11, 22, 33]);
});

test('buildBucketDeleteHandler does not call onStart when confirm returns false', async () => {
  let startCalled = false;

  const bucket = makeBucket([makeSession({ sessionId: 1 }), makeSession({ sessionId: 2 })]);

  const handler = buildBucketDeleteHandler({
    bucket,
    apiClient: { deleteSessions: async () => {} },
    confirm: () => false,
    onStart: () => {
      startCalled = true;
    },
    onSuccess: () => {},
    onError: () => {},
  });

  await handler();

  assert.equal(startCalled, false);
});

test('buildBucketDeleteHandler is a no-op when confirm returns false', async () => {
  let deleteCalled = false;
  let successCalled = false;

  const bucket = makeBucket([makeSession({ sessionId: 1 }), makeSession({ sessionId: 2 })]);

  const handler = buildBucketDeleteHandler({
    bucket,
    apiClient: {
      deleteSessions: async () => {
        deleteCalled = true;
      },
    },
    confirm: () => false,
    onSuccess: () => {
      successCalled = true;
    },
    onError: () => {},
  });

  await handler();

  assert.equal(deleteCalled, false);
  assert.equal(successCalled, false);
});

test('buildBucketDeleteHandler reports errors via onError without calling onSuccess', async () => {
  let errorMessage: string | null = null;
  let successCalled = false;

  const bucket = makeBucket([makeSession({ sessionId: 1 }), makeSession({ sessionId: 2 })]);

  const handler = buildBucketDeleteHandler({
    bucket,
    apiClient: {
      deleteSessions: async () => {
        throw new Error('boom');
      },
    },
    confirm: () => true,
    onSuccess: () => {
      successCalled = true;
    },
    onError: (message) => {
      errorMessage = message;
    },
  });

  await handler();

  assert.equal(errorMessage, 'boom');
  assert.equal(successCalled, false);
});

test('buildBucketDeleteHandler reports confirmation errors via onError', async () => {
  let errorMessage: string | null = null;
  let deleteCalled = false;

  const bucket = makeBucket([makeSession({ sessionId: 1 }), makeSession({ sessionId: 2 })]);

  const handler = buildBucketDeleteHandler({
    bucket,
    apiClient: {
      deleteSessions: async () => {
        deleteCalled = true;
      },
    },
    confirm: async () => {
      throw new Error('confirm failed');
    },
    onSuccess: () => {},
    onError: (message) => {
      errorMessage = message;
    },
  });

  await handler();

  assert.equal(errorMessage, 'confirm failed');
  assert.equal(deleteCalled, false);
});

test('buildBucketDeleteHandler falls back to a generic title when canonicalTitle is null', async () => {
  let seenTitle: string | null = null;

  const bucket = makeBucket([
    makeSession({ sessionId: 1, canonicalTitle: null }),
    makeSession({ sessionId: 2, canonicalTitle: null }),
  ]);

  const handler = buildBucketDeleteHandler({
    bucket,
    apiClient: { deleteSessions: async () => {} },
    confirm: (title) => {
      seenTitle = title;
      return false;
    },
    onSuccess: () => {},
    onError: () => {},
  });

  await handler();

  assert.equal(seenTitle, 'this episode');
});
