import assert from 'node:assert/strict';
import test from 'node:test';

import { buildCumulativePerAnime } from './anime-cumulative';

test('buildCumulativePerAnime carries each title forward to the last active day', () => {
  const result = buildCumulativePerAnime([
    { epochDay: 103, animeTitle: 'B', value: 4 },
    { epochDay: 100, animeTitle: 'A', value: 10 },
    { epochDay: 102, animeTitle: 'A', value: 5 },
    { epochDay: 102, animeTitle: 'A', value: 1 },
  ]);

  assert.deepEqual(
    result
      .filter((point) => point.animeTitle === 'A')
      .map((point) => [point.epochDay, point.value]),
    [
      [100, 10],
      [101, 10],
      [102, 16],
      [103, 16],
    ],
  );
  // A title starts at its own first active day, not the earliest day overall.
  assert.deepEqual(
    result
      .filter((point) => point.animeTitle === 'B')
      .map((point) => [point.epochDay, point.value]),
    [[103, 4]],
  );
});

test('buildCumulativePerAnime returns nothing for no points', () => {
  assert.deepEqual(buildCumulativePerAnime([]), []);
});
