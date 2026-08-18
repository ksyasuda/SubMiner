import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';

import { createCharacterDictionaryRuntimeService } from '../character-dictionary-runtime';
import { getSnapshotPath, writeSnapshot } from './cache';
import { CHARACTER_DICTIONARY_FORMAT_VERSION } from './constants';
import type { CharacterDictionarySnapshot } from './types';

const GRAPHQL_URL = 'https://graphql.anilist.co';
const PNG_1X1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nmX8AAAAASUVORK5CYII=',
  'base64',
);

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-dictionary-'));
}

function createSnapshotWithoutImages(): CharacterDictionarySnapshot {
  return {
    formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
    mediaId: 130298,
    mediaTitle: 'The Eminence in Shadow',
    entryCount: 1,
    updatedAt: 1_700_000_000_000,
    termEntries: [['アレクシア', 'あれくしあ', 'name primary', '', 75, ['Alexia'], 0, '']],
    images: [],
  };
}

test('generateForCurrentMedia refreshes same-version snapshots missing images when inline images are enabled', async () => {
  const userDataPath = makeTempDir();
  const outputDir = path.join(userDataPath, 'character-dictionaries');
  await writeSnapshot(getSnapshotPath(outputDir, 130298), createSnapshotWithoutImages());
  const originalFetch = globalThis.fetch;
  const fetchUrls: string[] = [];

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    fetchUrls.push(url);

    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
      };
      if (body.query?.includes('characters(page: $page')) {
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  english: 'The Eminence in Shadow',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'SUPPORTING',
                      node: {
                        id: 123,
                        description: 'Alexia Midgar.',
                        image: {
                          large: 'https://cdn.example.com/character-123.png',
                          medium: null,
                        },
                        name: {
                          full: 'Alexia Midgar',
                          native: 'アレクシア・ミドガル',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }
    }

    if (url === 'https://cdn.example.com/character-123.png') {
      return new Response(PNG_1X1, {
        status: 200,
        headers: { 'content-type': 'image/png' },
      });
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        season: null,
        episode: 5,
        source: 'fallback',
      }),
      getNameMatchImagesEnabled: () => true,
      now: () => 1_700_000_000_500,
    });

    const result = await runtime.generateForCurrentMedia();
    const refreshedSnapshot = JSON.parse(
      fs.readFileSync(getSnapshotPath(outputDir, 130298), 'utf8'),
    ) as CharacterDictionarySnapshot;

    assert.equal(result.fromCache, false);
    assert.ok(fetchUrls.includes(GRAPHQL_URL));
    assert.ok(refreshedSnapshot.images.some((image) => image.path === 'img/m130298-c123.png'));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia keeps failed MeCab name split refreshes retryable', async () => {
  const userDataPath = makeTempDir();
  const outputDir = path.join(userDataPath, 'character-dictionaries');
  await writeSnapshot(getSnapshotPath(outputDir, 130298), {
    ...createSnapshotWithoutImages(),
    nameSplitSource: 'heuristic',
  });
  const originalFetch = globalThis.fetch;

  let characterPageRequests = 0;
  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as { query?: string };
      if (body.query?.includes('characters(page: $page')) {
        characterPageRequests += 1;
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: { english: 'The Eminence in Shadow' },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'SUPPORTING',
                      node: {
                        id: 123,
                        description: 'Alexia Midgar.',
                        image: { large: null, medium: null },
                        name: {
                          first: 'Taro',
                          last: 'Yamada',
                          full: 'Taro Yamada',
                          native: '山田太郎',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }
    }
    throw new Error(`Unexpected fetch URL: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    let tokenizerCalls = 0;
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        season: null,
        episode: 5,
        source: 'fallback',
      }),
      getNameMatchImagesEnabled: () => false,
      tokenizeJapaneseName: async () => {
        tokenizerCalls += 1;
        return null;
      },
      getJapaneseNameTokenizerAvailable: () => true,
      now: () => 1_700_000_000_500,
    });

    const result = await runtime.generateForCurrentMedia();
    const refreshedSnapshot = JSON.parse(
      fs.readFileSync(getSnapshotPath(outputDir, 130298), 'utf8'),
    ) as CharacterDictionarySnapshot;

    assert.equal(result.fromCache, false);
    assert.equal(refreshedSnapshot.nameSplitSource, 'heuristic');

    const retriedResult = await runtime.generateForCurrentMedia();
    assert.equal(retriedResult.fromCache, false);
    assert.equal(characterPageRequests, 2);
    assert.equal(tokenizerCalls, 2);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia keeps mecab-split snapshots when MeCab is available', async () => {
  const userDataPath = makeTempDir();
  const outputDir = path.join(userDataPath, 'character-dictionaries');
  await writeSnapshot(getSnapshotPath(outputDir, 130298), {
    ...createSnapshotWithoutImages(),
    nameSplitSource: 'mecab',
  });
  const originalFetch = globalThis.fetch;

  globalThis.fetch = (async (input: string | URL | Request) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    throw new Error(`Unexpected fetch URL: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        season: null,
        episode: 5,
        source: 'fallback',
      }),
      getNameMatchImagesEnabled: () => false,
      tokenizeJapaneseName: async () => null,
      getJapaneseNameTokenizerAvailable: () => true,
      now: () => 1_700_000_000_500,
    });

    const result = await runtime.generateForCurrentMedia();

    assert.equal(result.fromCache, true);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia keeps heuristic-split snapshots while MeCab is unavailable', async () => {
  const userDataPath = makeTempDir();
  const outputDir = path.join(userDataPath, 'character-dictionaries');
  await writeSnapshot(getSnapshotPath(outputDir, 130298), {
    ...createSnapshotWithoutImages(),
    nameSplitSource: 'heuristic',
  });
  const originalFetch = globalThis.fetch;

  globalThis.fetch = (async (input: string | URL | Request) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    throw new Error(`Unexpected fetch URL: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        season: null,
        episode: 5,
        source: 'fallback',
      }),
      getNameMatchImagesEnabled: () => false,
      tokenizeJapaneseName: async () => null,
      getJapaneseNameTokenizerAvailable: () => false,
      now: () => 1_700_000_000_500,
    });

    const result = await runtime.generateForCurrentMedia();

    assert.equal(result.fromCache, true);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia keeps same-version snapshots without images when inline images are disabled', async () => {
  const userDataPath = makeTempDir();
  const outputDir = path.join(userDataPath, 'character-dictionaries');
  await writeSnapshot(getSnapshotPath(outputDir, 130298), createSnapshotWithoutImages());
  const originalFetch = globalThis.fetch;

  globalThis.fetch = (async (input: string | URL | Request) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    throw new Error(`Unexpected fetch URL: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        season: null,
        episode: 5,
        source: 'fallback',
      }),
      getNameMatchImagesEnabled: () => false,
      now: () => 1_700_000_000_500,
    });

    const result = await runtime.generateForCurrentMedia();

    assert.equal(result.fromCache, true);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('an unresolvable season is not cached as a normal AniList match', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  let searchCalls = 0;

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url !== GRAPHQL_URL) {
      return new Response(PNG_1X1, { status: 200, headers: { 'content-type': 'image/png' } });
    }

    const body = JSON.parse(String(init?.body ?? '{}')) as {
      query?: string;
      variables?: { search?: string; id?: number };
    };

    if (body.query?.includes('characters(page: $page')) {
      return new Response(
        JSON.stringify({
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
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      );
    }

    if (typeof body.variables?.id === 'number') {
      // No sequel edges: season 3 is unreachable from the season 1 anchor.
      return new Response(JSON.stringify({ data: { Media: { relations: { edges: [] } } } }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    }

    searchCalls += 1;
    return new Response(
      JSON.stringify({
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
      }),
      { status: 200, headers: { 'content-type': 'application/json' } },
    );
  }) as typeof globalThis.fetch;

  const warnings: string[] = [];
  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () =>
        '/anime/Oregairu/My Teen Romantic Comedy SNAFU (2013) - S03E01.mkv',
      getCurrentMediaTitle: () => null,
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'My Teen Romantic Comedy SNAFU',
        year: 2013,
        season: 3,
        episode: 1,
        source: 'guessit',
      }),
      now: () => 1_700_000_000_000,
      sleep: async () => {},
      logWarn: (message) => warnings.push(message),
    });

    await runtime.getOrCreateCurrentSnapshot();
    const searchesAfterFirst = searchCalls;
    await runtime.getOrCreateCurrentSnapshot();

    // Re-resolved rather than served from the resolution cache, and warned both times.
    assert.ok(searchCalls > searchesAfterFirst);
    assert.equal(warnings.filter((m) => /could not find season 3/i.test(m)).length, 2);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
