import assert from 'node:assert/strict';
import * as fs from 'fs';
import test, { type TestContext } from 'node:test';

import { getSnapshotPath, writeSnapshot } from './cache';
import { ANILIST_GRAPHQL_URL, CHARACTER_DICTIONARY_FORMAT_VERSION } from './constants';
import {
  EMINENCE_MEDIA_ID,
  EMINENCE_TITLE,
  PNG_1X1,
  createRuntime,
  installFakeAniList,
  installFetch,
  jsonResponse,
  type FakeAniListCharacterEdge,
} from './runtime-test-harness';
import type {
  CharacterDictionaryRuntimeDeps,
  CharacterDictionarySnapshot,
  NameSplitTokenizer,
} from './types';

function createSnapshotWithoutImages(
  overrides: Partial<CharacterDictionarySnapshot> = {},
): CharacterDictionarySnapshot {
  return {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: EMINENCE_MEDIA_ID,
    mediaTitle: 'The Eminence in Shadow',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [['アレクシア', 'あれくしあ', 'name primary', '', 75, ['Alexia'], 0, '']],
    images: [],
    ...overrides,
  };
}

/** Seeds a cached snapshot, then returns a runtime that will find it for the current episode. */
async function createRuntimeWithSnapshot(
  snapshot: CharacterDictionarySnapshot,
  overrides: Partial<CharacterDictionaryRuntimeDeps> = {},
) {
  const harness = createRuntime({ now: () => 1_700_000_000_500, ...overrides });
  await writeSnapshot(getSnapshotPath(harness.outputDir, EMINENCE_MEDIA_ID), snapshot);
  return harness;
}

function readStoredSnapshot(outputDir: string): CharacterDictionarySnapshot {
  return JSON.parse(
    fs.readFileSync(getSnapshotPath(outputDir, EMINENCE_MEDIA_ID), 'utf8'),
  ) as CharacterDictionarySnapshot;
}

function installCharacterPage(t: TestContext, edge: FakeAniListCharacterEdge, images?: string[]) {
  return installFakeAniList(t, {
    media: [
      { id: EMINENCE_MEDIA_ID, title: { english: EMINENCE_TITLE.english }, characters: [edge] },
    ],
    images,
  });
}

test('generateForCurrentMedia refreshes same-version snapshots missing images when inline images are enabled', async (t) => {
  const imageUrl = 'https://cdn.example.com/character-123.png';
  const anilist = installCharacterPage(
    t,
    {
      role: 'SUPPORTING',
      node: {
        id: 123,
        description: 'Alexia Midgar.',
        image: { large: imageUrl, medium: null },
        name: { full: 'Alexia Midgar', native: 'アレクシア・ミドガル' },
      },
    },
    [imageUrl],
  );
  const { runtime, outputDir } = await createRuntimeWithSnapshot(createSnapshotWithoutImages(), {
    getNameMatchImagesEnabled: () => true,
  });

  const result = await runtime.generateForCurrentMedia();

  assert.equal(result.fromCache, false);
  assert.equal(anilist.characterPageCount, 1);
  assert.ok(
    readStoredSnapshot(outputDir).images.some((image) => image.path === 'img/m130298-c123.png'),
  );
});

async function runNameSplitRefreshScenario(
  t: TestContext,
  tokenizeJapaneseName: NameSplitTokenizer,
): Promise<{
  characterPageRequests: number;
  firstResultFromCache: boolean;
  refreshedNameSplitSource: CharacterDictionarySnapshot['nameSplitSource'];
  secondResultFromCache: boolean;
}> {
  const anilist = installCharacterPage(t, {
    role: 'SUPPORTING',
    node: {
      id: 123,
      description: 'Alexia Midgar.',
      image: { large: null, medium: null },
      name: { first: 'Taro', last: 'Yamada', full: 'Taro Yamada', native: '山田太郎' },
    },
  });
  const { runtime, outputDir } = await createRuntimeWithSnapshot(
    createSnapshotWithoutImages({ nameSplitSource: 'heuristic' }),
    {
      getNameMatchImagesEnabled: () => false,
      tokenizeJapaneseName,
      getJapaneseNameTokenizerAvailable: () => true,
    },
  );

  const firstResult = await runtime.generateForCurrentMedia();
  const refreshedSnapshot = readStoredSnapshot(outputDir);
  const secondResult = await runtime.generateForCurrentMedia();

  return {
    characterPageRequests: anilist.characterPageCount,
    firstResultFromCache: firstResult.fromCache,
    refreshedNameSplitSource: refreshedSnapshot.nameSplitSource,
    secondResultFromCache: secondResult.fromCache,
  };
}

test('generateForCurrentMedia keeps failed MeCab name split refreshes retryable', async (t) => {
  let tokenizerCalls = 0;
  const result = await runNameSplitRefreshScenario(t, async () => {
    tokenizerCalls += 1;
    return null;
  });

  assert.equal(result.firstResultFromCache, false);
  assert.equal(result.refreshedNameSplitSource, 'heuristic');
  assert.equal(result.secondResultFromCache, false);
  assert.equal(result.characterPageRequests, 2);
  assert.equal(tokenizerCalls, 2);
});

test('generateForCurrentMedia caches completed MeCab refreshes with no resolved splits', async (t) => {
  let tokenizerCalls = 0;
  const result = await runNameSplitRefreshScenario(t, async () => {
    tokenizerCalls += 1;
    return [];
  });

  assert.equal(result.firstResultFromCache, false);
  assert.equal(result.refreshedNameSplitSource, 'mecab');
  assert.equal(result.secondResultFromCache, true);
  assert.equal(result.characterPageRequests, 1);
  assert.equal(tokenizerCalls, 1);
});

const keptSnapshotCases: Array<{
  name: string;
  snapshot: CharacterDictionarySnapshot;
  deps: Partial<CharacterDictionaryRuntimeDeps>;
}> = [
  {
    name: 'keeps mecab-split snapshots when MeCab is available',
    snapshot: createSnapshotWithoutImages({ nameSplitSource: 'mecab' }),
    deps: {
      getNameMatchImagesEnabled: () => false,
      tokenizeJapaneseName: async () => null,
      getJapaneseNameTokenizerAvailable: () => true,
    },
  },
  {
    name: 'keeps heuristic-split snapshots while MeCab is unavailable',
    snapshot: createSnapshotWithoutImages({ nameSplitSource: 'heuristic' }),
    deps: {
      getNameMatchImagesEnabled: () => false,
      tokenizeJapaneseName: async () => null,
      getJapaneseNameTokenizerAvailable: () => false,
    },
  },
  {
    name: 'keeps same-version snapshots without images when inline images are disabled',
    snapshot: createSnapshotWithoutImages(),
    deps: { getNameMatchImagesEnabled: () => false },
  },
];

for (const { name, snapshot, deps } of keptSnapshotCases) {
  test(`generateForCurrentMedia ${name}`, async (t) => {
    // No media or images: any AniList request throws.
    installFakeAniList(t);
    const { runtime } = await createRuntimeWithSnapshot(snapshot, deps);

    const result = await runtime.generateForCurrentMedia();

    assert.equal(result.fromCache, true);
  });
}

test('an unresolvable season is not cached as a normal AniList match', async (t) => {
  let searchCalls = 0;

  installFetch(t, (url, init) => {
    if (url !== ANILIST_GRAPHQL_URL) {
      return new Response(PNG_1X1, { status: 200, headers: { 'content-type': 'image/png' } });
    }

    const body = JSON.parse(String(init?.body ?? '{}')) as {
      query?: string;
      variables?: { search?: string; id?: number };
    };

    if (body.query?.includes('characters(page: $page')) {
      return jsonResponse({
        data: {
          Media: {
            title: { english: 'My Teen Romantic Comedy SNAFU' },
            characters: {
              pageInfo: { hasNextPage: false },
              edges: [
                {
                  role: 'MAIN',
                  node: { id: 1, name: { full: 'Hachiman Hikigaya', native: '比企谷八幡' } },
                },
              ],
            },
          },
        },
      });
    }

    if (typeof body.variables?.id === 'number') {
      // No sequel edges: season 3 is unreachable from the season 1 anchor.
      return jsonResponse({ data: { Media: { relations: { edges: [] } } } });
    }

    searchCalls += 1;
    return jsonResponse({
      data: {
        Page: {
          media: [
            {
              id: 14813,
              episodes: 13,
              format: 'TV',
              seasonYear: 2013,
              title: { romaji: null, english: 'My Teen Romantic Comedy SNAFU', native: null },
            },
          ],
        },
      },
    });
  });

  const warnings: string[] = [];
  const { runtime } = createRuntime({
    getCurrentMediaPath: () => '/anime/Oregairu/My Teen Romantic Comedy SNAFU (2013) - S03E01.mkv',
    getCurrentMediaTitle: () => null,
    guessAnilistMediaInfo: async () => ({
      title: 'My Teen Romantic Comedy SNAFU',
      year: 2013,
      season: 3,
      episode: 1,
      source: 'guessit',
    }),
    logWarn: (message) => warnings.push(message),
  });

  await runtime.getOrCreateCurrentSnapshot();
  const searchesAfterFirst = searchCalls;
  await runtime.getOrCreateCurrentSnapshot();

  // Re-resolved rather than served from the resolution cache, and warned both times.
  assert.ok(searchCalls > searchesAfterFirst);
  assert.equal(warnings.filter((m) => /could not find season 3/i.test(m)).length, 2);
});
