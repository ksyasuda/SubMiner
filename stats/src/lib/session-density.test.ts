import assert from 'node:assert/strict';
import test from 'node:test';
import { buildSessionDensity, pickDensityBucketMs } from './session-density';
import { EventType } from '../types/stats';

function sample(sampleMs: number, linesSeen: number, tokensSeen: number) {
  return { sampleMs, totalWatchedMs: 0, activeWatchedMs: 0, linesSeen, tokensSeen, cardsMined: 0 };
}

test('pickDensityBucketMs grows the bucket until the session fits in the bucket budget', () => {
  assert.equal(pickDensityBucketMs(5 * 60_000), 10_000);
  assert.equal(pickDensityBucketMs(25 * 60_000), 30_000);
  assert.equal(pickDensityBucketMs(3 * 3_600_000), 300_000);
});

test('pickDensityBucketMs scales past the largest candidate so very long spans stay in budget', () => {
  const spanMs = 48 * 3_600_000;
  const bucketMs = pickDensityBucketMs(spanMs);
  assert.ok(Math.ceil(spanMs / bucketMs) <= 60);
  assert.equal(bucketMs % 60_000, 0);
});

test('buildSessionDensity converts cumulative samples into per-bucket known/unknown deltas', () => {
  const density = buildSessionDensity(
    [sample(0, 0, 0), sample(5_000, 1, 10), sample(12_000, 2, 25), sample(25_000, 3, 40)],
    [
      { linesSeen: 1, knownWordsSeen: 4, totalWordsSeen: 8 },
      { linesSeen: 2, knownWordsSeen: 10, totalWordsSeen: 20 },
      { linesSeen: 3, knownWordsSeen: 20, totalWordsSeen: 32 },
    ],
    [
      { eventType: EventType.YOMITAN_LOOKUP, tsMs: 3_000, payload: null },
      { eventType: EventType.YOMITAN_LOOKUP, tsMs: 14_000, payload: null },
      { eventType: EventType.CARD_MINED, tsMs: 24_000, payload: '{"cardsMined":2}' },
      { eventType: EventType.PAUSE_START, tsMs: 24_500, payload: null },
    ],
  );

  assert.equal(density.bucketMs, 10_000);
  assert.equal(density.hasKnownWords, true);
  assert.deepEqual(
    density.buckets.map((b) => [b.knownWords, b.unknownWords, b.lookups, b.cards]),
    [
      [4, 4, 1, 0],
      [6, 6, 1, 0],
      [10, 2, 0, 2],
    ],
  );
});

test('buildSessionDensity falls back to raw token counts without known-word data', () => {
  const density = buildSessionDensity(
    [sample(0, 0, 0), sample(9_000, 1, 7), sample(15_000, 2, 12)],
    [],
    [],
  );

  assert.equal(density.hasKnownWords, false);
  assert.deepEqual(
    density.buckets.map((b) => [b.words, b.knownWords]),
    [
      [7, 0],
      [5, 0],
    ],
  );
});

test('buildSessionDensity returns no buckets for an empty timeline', () => {
  assert.deepEqual(buildSessionDensity([], [], []).buckets, []);
});

test('buildSessionDensity keeps the final sample when the span ends exactly on a bucket edge', () => {
  const density = buildSessionDensity([sample(0, 0, 0), sample(300_000, 1, 42)], [], []);

  assert.equal(density.bucketMs, 10_000);
  assert.equal(density.buckets.length, 30);
  assert.equal(density.buckets.at(-1)!.words, 42);
  assert.equal(
    density.buckets.reduce((sum, b) => sum + b.words, 0),
    42,
  );
});
