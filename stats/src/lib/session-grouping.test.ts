import assert from 'node:assert/strict';
import test from 'node:test';
import type { SessionSummary } from '../types/stats';
import { groupSessionsByVideo } from './session-grouping';

function makeSession(overrides: Partial<SessionSummary> & { sessionId: number }): SessionSummary {
  const { sessionId, ...rest } = overrides;
  return {
    sessionId,
    canonicalTitle: null,
    videoId: null,
    animeId: null,
    animeTitle: null,
    startedAtMs: 1000,
    endedAtMs: null,
    totalWatchedMs: 0,
    activeWatchedMs: 0,
    linesSeen: 0,
    tokensSeen: 0,
    cardsMined: 0,
    lookupCount: 0,
    lookupHits: 0,
    yomitanLookupCount: 0,
    knownWordsSeen: 0,
    knownWordRate: 0,
    ...rest,
  };
}

test('empty input returns empty array', () => {
  assert.deepEqual(groupSessionsByVideo([]), []);
});

test('two unique videoIds produce 2 singleton buckets', () => {
  const sessions = [
    makeSession({
      sessionId: 1,
      videoId: 10,
      startedAtMs: 1000,
      activeWatchedMs: 100,
      cardsMined: 2,
    }),
    makeSession({
      sessionId: 2,
      videoId: 20,
      startedAtMs: 2000,
      activeWatchedMs: 200,
      cardsMined: 3,
    }),
  ];
  const buckets = groupSessionsByVideo(sessions);
  assert.equal(buckets.length, 2);
  const keys = buckets.map((b) => b.key).sort();
  assert.deepEqual(keys, ['v-10', 'v-20']);
  for (const bucket of buckets) {
    assert.equal(bucket.sessions.length, 1);
  }
});

test('two sessions sharing a videoId collapse into 1 bucket with summed totals and most-recent representative', () => {
  const older = makeSession({
    sessionId: 1,
    videoId: 42,
    startedAtMs: 1000,
    activeWatchedMs: 300,
    cardsMined: 5,
  });
  const newer = makeSession({
    sessionId: 2,
    videoId: 42,
    startedAtMs: 9000,
    activeWatchedMs: 500,
    cardsMined: 7,
  });
  const buckets = groupSessionsByVideo([older, newer]);
  assert.equal(buckets.length, 1);
  const [bucket] = buckets;
  assert.equal(bucket!.key, 'v-42');
  assert.equal(bucket!.videoId, 42);
  assert.equal(bucket!.sessions.length, 2);
  assert.equal(bucket!.totalActiveMs, 800);
  assert.equal(bucket!.totalCardsMined, 12);
  assert.equal(bucket!.representativeSession.sessionId, 2); // most recent (highest startedAtMs)
});

test('sessions with null videoId become singleton buckets keyed by sessionId', () => {
  const s1 = makeSession({ sessionId: 101, videoId: null, activeWatchedMs: 50, cardsMined: 1 });
  const s2 = makeSession({ sessionId: 202, videoId: null, activeWatchedMs: 75, cardsMined: 2 });
  const buckets = groupSessionsByVideo([s1, s2]);
  assert.equal(buckets.length, 2);
  const keys = buckets.map((b) => b.key).sort();
  assert.deepEqual(keys, ['s-101', 's-202']);
  for (const bucket of buckets) {
    assert.equal(bucket.videoId, null);
    assert.equal(bucket.sessions.length, 1);
  }
});
