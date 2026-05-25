import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';

import { createCharacterDictionaryRuntimeService } from '../character-dictionary-runtime';
import { buildCharacterDictionarySeriesKey } from './manual-selection';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-dictionary-'));
}

test('getManualSelectionSnapshot waits for explicit search text before fetching candidates', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  const searchTerms: string[] = [];

  globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
    const body = JSON.parse(String(init?.body ?? '{}')) as {
      variables?: { search?: string };
    };
    searchTerms.push(String(body.variables?.search ?? ''));
    return new Response(
      JSON.stringify({
        data: {
          Page: {
            media: [
              {
                id: 154587,
                episodes: 28,
                title: {
                  romaji: 'Sousou no Frieren',
                  english: 'Frieren: Beyond Journey’s End',
                  native: '葬送のフリーレン',
                },
              },
            ],
          },
        },
      }),
      {
        status: 200,
        headers: { 'content-type': 'application/json' },
      },
    );
  }) as typeof globalThis.fetch;

  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/[SubsPlease] Kage no Jitsuryokusha - 05.mkv',
      getCurrentMediaTitle: () => '[SubsPlease] Kage no Jitsuryokusha - 05.mkv',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'Kage no Jitsuryokusha ni Naritakute!',
        season: null,
        episode: 5,
        source: 'guessit',
      }),
      now: () => 1_700_000_000_000,
    });

    const initial = await runtime.getManualSelectionSnapshot(undefined, '');
    assert.equal(initial.guessTitle, 'Kage no Jitsuryokusha ni Naritakute!');
    assert.deepEqual(initial.candidates, []);
    assert.deepEqual(searchTerms, []);

    const searched = await runtime.getManualSelectionSnapshot(undefined, 'Frieren');
    assert.deepEqual(searchTerms, ['Frieren']);
    assert.deepEqual(searched.candidates, [
      { id: 154587, title: 'Frieren: Beyond Journey’s End', episodes: 28 },
    ]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getManualSelectionSnapshot hydrates override episode count from searched candidates', async () => {
  const userDataPath = makeTempDir();
  const overrideSeriesKey = buildCharacterDictionarySeriesKey({
    mediaPath: '/tmp/KonoSuba - 01.mkv',
    mediaTitle: 'KonoSuba - 01.mkv',
    guess: {
      title: "KonoSuba - God's blessing on this wonderful world!",
      year: 2016,
      season: null,
      episode: 1,
      source: 'guessit',
    },
  });
  const overrideDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(overrideDir, { recursive: true });
  fs.writeFileSync(
    path.join(overrideDir, 'anilist-overrides.json'),
    JSON.stringify({
      overrides: [
        {
          seriesKey: overrideSeriesKey,
          mediaId: 21202,
          mediaTitle: "KONOSUBA -God's blessing on this wonderful world!",
          staleMediaIds: [],
        },
      ],
    }),
    'utf8',
  );
  const originalFetch = globalThis.fetch;

  globalThis.fetch = (async (_input: string | URL | Request) => {
    return new Response(
      JSON.stringify({
        data: {
          Page: {
            media: [
              {
                id: 21202,
                episodes: 10,
                title: {
                  romaji: 'Kono Subarashii Sekai ni Shukufuku wo!',
                  english: "KONOSUBA -God's blessing on this wonderful world!",
                  native: 'この素晴らしい世界に祝福を！',
                },
              },
            ],
          },
        },
      }),
      {
        status: 200,
        headers: { 'content-type': 'application/json' },
      },
    );
  }) as typeof globalThis.fetch;

  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/KonoSuba - 01.mkv',
      getCurrentMediaTitle: () => 'KonoSuba - 01.mkv',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: "KonoSuba - God's blessing on this wonderful world!",
        year: 2016,
        season: null,
        episode: 1,
        source: 'guessit',
      }),
      now: () => 1_700_000_000_000,
    });

    const snapshot = await runtime.getManualSelectionSnapshot(undefined, 'KonoSuba');

    assert.deepEqual(snapshot.override, {
      id: 21202,
      title: "KONOSUBA -God's blessing on this wonderful world!",
      episodes: 10,
    });
  } finally {
    globalThis.fetch = originalFetch;
  }
});
