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
  '/anime/ReZERO/Season 1/Re - ZERO, Starting Life in Another World (2016) - S01E01 - - The End of the Beginning and the Beginning of the End [v2 Bluray-1080p Proper][10bit][x265][FLAC 2.0][EN+JA]-SCY.mkv';
const REZERO_EP2 =
  '/anime/ReZERO/Season 1/Re - ZERO, Starting Life in Another World (2016) - S01E02 - Reunion with the Witch [Bluray-1080p][x265][JA]-SCY.mkv';
const REZERO_S2_EP1 =
  '/anime/ReZERO/Season 2/Re - ZERO, Starting Life in Another World (2016) - S02E01 - Each Ones Promise [Bluray-1080p][x265][JA]-SCY.mkv';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-manual-selection-'));
}

test('buildCharacterDictionarySeriesKey scopes guessit title and year by media directory', () => {
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

  assert.equal(key, 'anime-rezero-season-1--re-zero-starting-life-in-another-world-2016');
});

test('manual selection store persists overrides and matches later episodes in the same directory', async () => {
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

test('manual selection store applies a season-directory override when later episode guesses differ', async () => {
  const userDataPath = makeTempDir();
  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });
  const directory = '/Volumes/jellyfin/anime/toaru-kagaku-no-railgun/Season-2';
  const selectedKey = buildCharacterDictionarySeriesKey({
    mediaPath: `${directory}/A Certain Scientific Railgun S - 10 - Critical.mkv`,
    mediaTitle: 'A Certain Scientific Railgun S - 10 - Critical.mkv',
    guess: {
      title: 'Critical',
      season: null,
      episode: 10,
      source: 'guessit',
    },
  });
  await store.setOverride({
    seriesKey: selectedKey,
    mediaId: 16049,
    mediaTitle: 'A Certain Scientific Railgun S',
    staleMediaIds: [1057],
  });

  const laterEpisodeKey = buildCharacterDictionarySeriesKey({
    mediaPath: `${directory}/A Certain Scientific Railgun S - 16 - Sisters.mkv`,
    mediaTitle: 'A Certain Scientific Railgun S - 16 - Sisters.mkv',
    guess: {
      title: 'Sisters',
      season: null,
      episode: 16,
      source: 'guessit',
    },
  });

  assert.deepEqual(await store.getOverride(laterEpisodeKey), {
    seriesKey: selectedKey,
    mediaId: 16049,
    mediaTitle: 'A Certain Scientific Railgun S',
    staleMediaIds: [1057],
  });
});

test('manual selection store replaces older guessed keys for the same season directory', async () => {
  const userDataPath = makeTempDir();
  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });
  const directoryKey = 'volumes-jellyfin-anime-toaru-kagaku-no-railgun-season-2';
  const firstKey = `${directoryKey}--critical`;
  const laterKey = `${directoryKey}--sisters`;
  await store.setOverride({
    seriesKey: firstKey,
    mediaId: 1057,
    mediaTitle: 'Ippatsu Kiki Musume',
    staleMediaIds: [],
  });
  await store.setOverride({
    seriesKey: laterKey,
    mediaId: 16049,
    mediaTitle: 'A Certain Scientific Railgun S',
    staleMediaIds: [1057],
  });

  assert.deepEqual(await store.getOverride(firstKey), {
    seriesKey: laterKey,
    mediaId: 16049,
    mediaTitle: 'A Certain Scientific Railgun S',
    staleMediaIds: [1057],
  });
});

test('manual selection store resolves legacy unscoped override keys', async () => {
  const userDataPath = makeTempDir();
  const overrideDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(overrideDir, { recursive: true });
  fs.writeFileSync(
    path.join(overrideDir, 'anilist-overrides.json'),
    JSON.stringify({
      overrides: [
        {
          seriesKey: 're-zero-starting-life-in-another-world-2016',
          mediaId: 21355,
          mediaTitle: 'Re:ZERO -Starting Life in Another World-',
          staleMediaIds: [10607],
        },
      ],
    }),
    'utf8',
  );

  const scopedKey = buildCharacterDictionarySeriesKey({
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

  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });

  assert.deepEqual(await store.getOverride(scopedKey), {
    seriesKey: 're-zero-starting-life-in-another-world-2016',
    mediaId: 21355,
    mediaTitle: 'Re:ZERO -Starting Life in Another World-',
    staleMediaIds: [10607],
  });
});

test('manual selection store prefers exact scoped override over legacy fallback', async () => {
  const userDataPath = makeTempDir();
  const overrideDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(overrideDir, { recursive: true });
  const scopedKey = buildCharacterDictionarySeriesKey({
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
  fs.writeFileSync(
    path.join(overrideDir, 'anilist-overrides.json'),
    JSON.stringify({
      overrides: [
        {
          seriesKey: 're-zero-starting-life-in-another-world-2016',
          mediaId: 10607,
          mediaTitle: 'Legacy Re:ZERO',
          staleMediaIds: [],
        },
        {
          seriesKey: scopedKey,
          mediaId: 21355,
          mediaTitle: 'Re:ZERO -Starting Life in Another World-',
          staleMediaIds: [10607],
        },
      ],
    }),
    'utf8',
  );

  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });

  assert.deepEqual(await store.getOverride(scopedKey), {
    seriesKey: scopedKey,
    mediaId: 21355,
    mediaTitle: 'Re:ZERO -Starting Life in Another World-',
    staleMediaIds: [10607],
  });
});

test('manual selection store prefers same-directory override over legacy fallback', async () => {
  const userDataPath = makeTempDir();
  const overrideDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(overrideDir, { recursive: true });
  const directoryScope = 'volumes-jellyfin-anime-toaru-kagaku-no-railgun-season-2';
  fs.writeFileSync(
    path.join(overrideDir, 'anilist-overrides.json'),
    JSON.stringify({
      overrides: [
        {
          seriesKey: 'sisters',
          mediaId: 1057,
          mediaTitle: 'Legacy fallback',
          staleMediaIds: [],
        },
        {
          seriesKey: `${directoryScope}--critical`,
          mediaId: 16049,
          mediaTitle: 'A Certain Scientific Railgun S',
          staleMediaIds: [1057],
        },
      ],
    }),
    'utf8',
  );

  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });

  assert.deepEqual(await store.getOverride(`${directoryScope}--sisters`), {
    seriesKey: `${directoryScope}--critical`,
    mediaId: 16049,
    mediaTitle: 'A Certain Scientific Railgun S',
    staleMediaIds: [1057],
  });
});

test('manual selection store keeps overrides separate for different season directories', async () => {
  const userDataPath = makeTempDir();
  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });
  const firstSeasonKey = buildCharacterDictionarySeriesKey({
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
    seriesKey: firstSeasonKey,
    mediaId: 21355,
    mediaTitle: 'Re:ZERO -Starting Life in Another World-',
    staleMediaIds: [],
  });

  const secondSeasonKey = buildCharacterDictionarySeriesKey({
    mediaPath: REZERO_S2_EP1,
    mediaTitle: null,
    guess: {
      title: 'Re ZERO, Starting Life in Another World',
      alternativeTitle: 'ZERO, Starting Life in Another World',
      year: 2016,
      season: 2,
      episode: 1,
      source: 'guessit',
    },
  });

  assert.notEqual(secondSeasonKey, firstSeasonKey);
  assert.equal(await store.getOverride(secondSeasonKey), null);
});

test('buildCharacterDictionarySeriesKey tags seasons past the first', () => {
  const base = {
    title: 'My Teen Romantic Comedy SNAFU',
    year: 2013,
    episode: 1,
    source: 'guessit' as const,
  };
  const seasonOne = buildCharacterDictionarySeriesKey({
    mediaPath: '/anime/Oregairu/My Teen Romantic Comedy SNAFU (2013) - S01E01.mkv',
    mediaTitle: null,
    guess: { ...base, season: 1 },
  });
  const seasonThree = buildCharacterDictionarySeriesKey({
    mediaPath: '/anime/Oregairu/My Teen Romantic Comedy SNAFU (2013) - S03E01.mkv',
    mediaTitle: null,
    guess: { ...base, season: 3 },
  });

  // Season 1 keeps the pre-existing key shape so cached snapshots stay valid.
  assert.equal(seasonOne, 'anime-oregairu--my-teen-romantic-comedy-snafu-2013');
  assert.equal(seasonThree, 'anime-oregairu--my-teen-romantic-comedy-snafu-s3-2013');
});

test('manual selection store keeps seasons apart inside one flat directory', async () => {
  const userDataPath = makeTempDir();
  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });
  const base = {
    title: 'My Teen Romantic Comedy SNAFU',
    year: 2013,
    episode: 1,
    source: 'guessit' as const,
  };
  const seasonOneKey = buildCharacterDictionarySeriesKey({
    mediaPath: '/anime/Oregairu/My Teen Romantic Comedy SNAFU (2013) - S01E01.mkv',
    mediaTitle: null,
    guess: { ...base, season: 1 },
  });
  const seasonThreeKey = buildCharacterDictionarySeriesKey({
    mediaPath: '/anime/Oregairu/My Teen Romantic Comedy SNAFU (2013) - S03E01.mkv',
    mediaTitle: null,
    guess: { ...base, season: 3 },
  });

  await store.setOverride({
    seriesKey: seasonOneKey,
    mediaId: 14813,
    mediaTitle: 'My Teen Romantic Comedy SNAFU',
    staleMediaIds: [],
  });

  // Same directory, different season: the season 1 override must not leak across.
  assert.equal(await store.getOverride(seasonThreeKey), null);

  await store.setOverride({
    seriesKey: seasonThreeKey,
    mediaId: 108489,
    mediaTitle: 'My Teen Romantic Comedy SNAFU Climax!',
    staleMediaIds: [],
  });

  assert.equal((await store.getOverride(seasonOneKey))?.mediaId, 14813);
  assert.equal((await store.getOverride(seasonThreeKey))?.mediaId, 108489);
});

test('override scope uses the recorded season, not season-like text in the title', async () => {
  const userDataPath = makeTempDir();
  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });
  // A title that itself normalizes to a trailing "-s2" would be misparsed as season 2.
  const trickyTitleKey = buildCharacterDictionarySeriesKey({
    mediaPath: '/anime/Mixed/Some Show S2 - S01E01.mkv',
    mediaTitle: null,
    guess: { title: 'Some Show S2', season: 1, episode: 1, source: 'guessit' },
  });
  const realSeasonTwoKey = buildCharacterDictionarySeriesKey({
    mediaPath: '/anime/Mixed/Other Show - S02E01.mkv',
    mediaTitle: null,
    guess: { title: 'Other Show', season: 2, episode: 1, source: 'guessit' },
  });

  await store.setOverride({
    seriesKey: trickyTitleKey,
    mediaId: 111,
    mediaTitle: 'Some Show S2',
    staleMediaIds: [],
    season: 1,
  });
  await store.setOverride({
    seriesKey: realSeasonTwoKey,
    mediaId: 222,
    mediaTitle: 'Other Show 2',
    staleMediaIds: [],
    season: 2,
  });

  // Same directory, genuinely different seasons: neither override may replace the other.
  assert.equal((await store.getOverride(trickyTitleKey, 1))?.mediaId, 111);
  assert.equal((await store.getOverride(realSeasonTwoKey, 2))?.mediaId, 222);
});

test('override records without a stored season still resolve via the key', async () => {
  const userDataPath = makeTempDir();
  const store = createCharacterDictionaryManualSelectionStore({ userDataPath });
  const legacyKey = buildCharacterDictionarySeriesKey({
    mediaPath: REZERO_EP1,
    mediaTitle: null,
    guess: {
      title: 'Re ZERO, Starting Life in Another World',
      year: 2016,
      season: 1,
      episode: 1,
      source: 'guessit',
    },
  });
  const overridesPath = path.join(userDataPath, 'character-dictionaries', 'anilist-overrides.json');
  fs.mkdirSync(path.dirname(overridesPath), { recursive: true });
  fs.writeFileSync(
    overridesPath,
    JSON.stringify({
      overrides: [
        { seriesKey: legacyKey, mediaId: 21355, mediaTitle: 'Re:ZERO', staleMediaIds: [] },
      ],
    }),
    'utf8',
  );

  assert.equal((await store.getOverride(legacyKey, 1))?.mediaId, 21355);
});
