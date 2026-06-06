import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildCoverImageRequestKey,
  collectSessionCoverRequests,
  getCoverImageKey,
} from './cover-images';
import type { SessionSummary } from '../types/stats';

function makeSession(overrides: Partial<SessionSummary> & { sessionId: number }): SessionSummary {
  const { sessionId, ...rest } = overrides;
  return {
    sessionId,
    canonicalTitle: null,
    videoId: null,
    animeId: null,
    animeTitle: null,
    startedAtMs: 0,
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

test('collectSessionCoverRequests dedupes anime ids and only requests media for ungrouped sessions', () => {
  const requests = collectSessionCoverRequests([
    makeSession({ sessionId: 1, animeId: 10, videoId: 100 }),
    makeSession({ sessionId: 2, animeId: 10, videoId: 101 }),
    makeSession({ sessionId: 3, animeId: null, videoId: 200 }),
    makeSession({ sessionId: 4, animeId: null, videoId: 200 }),
  ]);

  assert.deepEqual(requests, { animeIds: [10], videoIds: [200] });
});

test('getCoverImageKey separates anime and media ids', () => {
  assert.equal(getCoverImageKey('anime', 1), 'anime:1');
  assert.equal(getCoverImageKey('media', 1), 'media:1');
});

test('buildCoverImageRequestKey changes when callers force a cover refresh', () => {
  assert.notEqual(buildCoverImageRequestKey([10], [], 0), buildCoverImageRequestKey([10], [], 1));
});
