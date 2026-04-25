import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import {
  buildCharacterDictionarySeriesKey,
  createCharacterDictionaryManualSelectionStore,
} from './manual-selection';

const REZERO_EP1 =
  '/anime/Re - ZERO, Starting Life in Another World (2016) - S01E01 - - The End of the Beginning and the Beginning of the End [v2 Bluray-1080p Proper][10bit][x265][FLAC 2.0][EN+JA]-SCY.mkv';
const REZERO_EP2 =
  '/anime/Re - ZERO, Starting Life in Another World (2016) - S01E02 - Reunion with the Witch [Bluray-1080p][x265][JA]-SCY.mkv';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-manual-selection-'));
}

test('buildCharacterDictionarySeriesKey uses guessit title, alternative title, and year for Re ZERO series scope', () => {
  const key = buildCharacterDictionarySeriesKey({
    mediaPath: REZERO_EP1,
    mediaTitle: null,
    guess: {
      title: 'Re ZERO, Starting Life in Another World',
      alternativeTitle: 'ZERO, Starting Life in Another World',
      year: 2016,
      season: 1,
      episode: 1,
      source: 'guessit',
    },
  });

  assert.equal(key, 're-zero-starting-life-in-another-world-2016');
});

test('manual selection store persists overrides and matches later episodes in the same series', async () => {
  const userDataPath = makeTempDir();
  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });
  const firstKey = buildCharacterDictionarySeriesKey({
    mediaPath: REZERO_EP1,
    mediaTitle: null,
    guess: {
      title: 'Re ZERO, Starting Life in Another World',
      alternativeTitle: 'ZERO, Starting Life in Another World',
      year: 2016,
      season: 1,
      episode: 1,
      source: 'guessit',
    },
  });
  await store.setOverride({
    seriesKey: firstKey,
    mediaId: 21355,
    mediaTitle: 'Re:ZERO -Starting Life in Another World-',
    staleMediaIds: [10607],
  });

  const reloaded = createCharacterDictionaryManualSelectionStore({ userDataPath });
  const secondKey = buildCharacterDictionarySeriesKey({
    mediaPath: REZERO_EP2,
    mediaTitle: null,
    guess: {
      title: 'Re ZERO, Starting Life in Another World',
      alternativeTitle: 'ZERO, Starting Life in Another World',
      year: 2016,
      season: 1,
      episode: 2,
      source: 'guessit',
    },
  });

  assert.equal(secondKey, firstKey);
  assert.deepEqual(await reloaded.getOverride(secondKey), {
    seriesKey: firstKey,
    mediaId: 21355,
    mediaTitle: 'Re:ZERO -Starting Life in Another World-',
    staleMediaIds: [10607],
  });
});
