import assert from 'node:assert/strict';
import test from 'node:test';

import type { PerAnimeDataPoint } from './StackedTrendChart';
import {
  buildAnimeVisibilityOptions,
  filterHiddenAnimeData,
  pruneHiddenAnime,
} from './anime-visibility';

const SAMPLE_POINTS: PerAnimeDataPoint[] = [
  { epochDay: 1, animeTitle: 'KonoSuba', value: 5 },
  { epochDay: 2, animeTitle: 'KonoSuba', value: 10 },
  { epochDay: 1, animeTitle: 'Little Witch Academia', value: 6 },
  { epochDay: 1, animeTitle: 'Trapped in a Dating Sim', value: 20 },
];

test('buildAnimeVisibilityOptions sorts anime by combined contribution', () => {
  const titles = buildAnimeVisibilityOptions([
    SAMPLE_POINTS,
    [
      { epochDay: 1, animeTitle: 'Little Witch Academia', value: 8 },
      { epochDay: 1, animeTitle: 'KonoSuba', value: 1 },
    ],
  ]);

  assert.deepEqual(titles, ['Trapped in a Dating Sim', 'KonoSuba', 'Little Witch Academia']);
});

test('filterHiddenAnimeData removes globally hidden anime from chart data', () => {
  const filtered = filterHiddenAnimeData(SAMPLE_POINTS, new Set(['KonoSuba']));

  assert.equal(
    filtered.some((point) => point.animeTitle === 'KonoSuba'),
    false,
  );
  assert.equal(filtered.length, 2);
});

test('pruneHiddenAnime drops titles that are no longer available', () => {
  const hidden = pruneHiddenAnime(new Set(['KonoSuba', 'Ghost in the Shell']), [
    'KonoSuba',
    'Little Witch Academia',
  ]);

  assert.deepEqual([...hidden], ['KonoSuba']);
});
