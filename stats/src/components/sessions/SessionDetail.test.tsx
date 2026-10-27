import assert from 'node:assert/strict';
import test from 'node:test';
import { buildSessionStatTiles, pickDensityTickIndexes } from './SessionDetail';
import type { SessionSummary } from '../../types/stats';

const session: SessionSummary = {
  sessionId: 1,
  canonicalTitle: 'Terrace House S01E08',
  videoId: 1,
  animeId: null,
  animeTitle: null,
  startedAtMs: 0,
  endedAtMs: 1_860_000,
  totalWatchedMs: 1_860_000,
  activeWatchedMs: 1_500_000,
  linesSeen: 400,
  tokensSeen: 1_684,
  cardsMined: 0,
  lookupCount: 0,
  lookupHits: 0,
  yomitanLookupCount: 18,
  knownWordsSeen: 673,
  knownWordRate: 40,
};

test('buildSessionStatTiles derives pace, known rate, lookup rate, and paused time', () => {
  const tiles = Object.fromEntries(buildSessionStatTiles(session).map((t) => [t.key, t]));
  assert.equal(tiles.pace!.value, '67 words/min');
  assert.equal(tiles.pace!.detail, '16.0 lines/min');
  assert.equal(tiles.known!.value, '40%');
  assert.equal(tiles.lookups!.detail, '1.1 lookups per 100 words');
  assert.equal(tiles.paused!.value, '6m');
  assert.equal(tiles.paused!.detail, '19% of 31m wall time');
});

test('buildSessionStatTiles shows placeholders when there is no active time or known data', () => {
  const tiles = Object.fromEntries(
    buildSessionStatTiles({
      ...session,
      activeWatchedMs: 0,
      totalWatchedMs: 0,
      knownWordsSeen: 0,
      knownWordRate: 0,
    }).map((t) => [t.key, t]),
  );
  assert.equal(tiles.pace!.value, '—');
  assert.equal(tiles.known!.value, '—');
  assert.equal(tiles.paused!.value, 'none');
});

test('pickDensityTickIndexes spaces about five labels across the buckets', () => {
  assert.deepEqual(pickDensityTickIndexes(3), [0, 1, 2]);
  assert.deepEqual(pickDensityTickIndexes(50), [0, 10, 20, 30, 40]);
  assert.deepEqual(pickDensityTickIndexes(0), []);
});
