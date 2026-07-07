import assert from 'node:assert/strict';
import test from 'node:test';

import type { PerAnimeDataPoint } from './StackedTrendChart';
import {
  buildAnimeVisibilityOptions,
  filterHiddenAnimeData,
  loadHiddenTitles,
  loadMaxTitles,
  loadMaxTitlesMode,
  pruneHiddenAnime,
  saveHiddenTitles,
  saveMaxTitles,
  saveMaxTitlesMode,
} from './anime-visibility';

function installLocalStorage(initial: Record<string, string> = {}) {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const values = new Map(Object.entries(initial));
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
      removeItem: (key: string) => values.delete(key),
    },
  });
  return {
    values,
    restore: () => {
      if (previous) {
        Object.defineProperty(globalThis, 'localStorage', previous);
      } else {
        delete (globalThis as { localStorage?: unknown }).localStorage;
      }
    },
  };
}

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

test('hidden titles round-trip through localStorage', () => {
  const { restore } = installLocalStorage();
  try {
    saveHiddenTitles(new Set(['KonoSuba', 'One Piece']));
    assert.deepEqual([...loadHiddenTitles()], ['KonoSuba', 'One Piece']);
  } finally {
    restore();
  }
});

test('loadHiddenTitles tolerates missing or malformed stored values', () => {
  const { restore } = installLocalStorage({
    'subminer-stats-trends-hidden-titles': '{"not":"an array"}',
  });
  try {
    assert.deepEqual([...loadHiddenTitles()], []);
  } finally {
    restore();
  }
});

test('max titles preference round-trips and clears when set to null', () => {
  const { values, restore } = installLocalStorage();
  try {
    saveMaxTitles(7);
    assert.equal(loadMaxTitles(), 7);
    saveMaxTitles(null);
    assert.equal(loadMaxTitles(), null);
    assert.equal(values.has('subminer-stats-trends-max-titles'), false);
  } finally {
    restore();
  }
});

test('loadMaxTitles rejects unsupported stored values', () => {
  const { restore } = installLocalStorage({ 'subminer-stats-trends-max-titles': '8' });
  try {
    assert.equal(loadMaxTitles(), null);
  } finally {
    restore();
  }
});

test('loadMaxTitles rejects non-positive or non-numeric stored values', () => {
  for (const storedValue of ['0', 'banana']) {
    const { restore } = installLocalStorage({ 'subminer-stats-trends-max-titles': storedValue });
    try {
      assert.equal(loadMaxTitles(), null);
    } finally {
      restore();
    }
  }
});

test('max titles mode round-trips and defaults to total', () => {
  const { restore } = installLocalStorage();
  try {
    assert.equal(loadMaxTitlesMode(), 'total');
    saveMaxTitlesMode('recent');
    assert.equal(loadMaxTitlesMode(), 'recent');
    saveMaxTitlesMode('total');
    assert.equal(loadMaxTitlesMode(), 'total');
  } finally {
    restore();
  }
});

test('loadMaxTitlesMode falls back to total for unknown stored values', () => {
  const { restore } = installLocalStorage({ 'subminer-stats-trends-max-titles-mode': 'sideways' });
  try {
    assert.equal(loadMaxTitlesMode(), 'total');
  } finally {
    restore();
  }
});
