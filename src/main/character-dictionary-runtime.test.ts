import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { createCharacterDictionaryRuntimeService } from './character-dictionary-runtime';

const GRAPHQL_URL = 'https://graphql.anilist.co';
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const PNG_1X1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nmX8AAAAASUVORK5CYII=',
  'base64',
);

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-dictionary-'));
}

function readStoredZipEntry(zipPath: string, entryName: string): Buffer {
  const archive = fs.readFileSync(zipPath);
  let offset = 0;

  while (offset + 4 <= archive.length) {
    const signature = archive.readUInt32LE(offset);
    if (
      signature === CENTRAL_DIRECTORY_SIGNATURE ||
      signature === END_OF_CENTRAL_DIRECTORY_SIGNATURE
    ) {
      break;
    }
    if (signature !== LOCAL_FILE_HEADER_SIGNATURE) {
      throw new Error(`Unexpected ZIP signature 0x${signature.toString(16)} at offset ${offset}`);
    }

    const compressionMethod = archive.readUInt16LE(offset + 8);
    assert.equal(compressionMethod, 0, 'expected stored ZIP entry');
    const compressedSize = archive.readUInt32LE(offset + 18);
    const fileNameLength = archive.readUInt16LE(offset + 26);
    const extraFieldLength = archive.readUInt16LE(offset + 28);
    const fileNameStart = offset + 30;
    const fileNameEnd = fileNameStart + fileNameLength;
    const fileName = archive.subarray(fileNameStart, fileNameEnd).toString('utf8');
    const dataStart = fileNameEnd + extraFieldLength;
    const dataEnd = dataStart + compressedSize;

    if (fileName === entryName) {
      return archive.subarray(dataStart, dataEnd);
    }

    offset = dataEnd;
  }

  throw new Error(`ZIP entry not found: ${entryName}`);
}

test('generateForCurrentMedia emits structured-content glossary so image stays with text', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
        variables?: Record<string, unknown>;
      };

      if (body.query?.includes('Page(perPage: 10)')) {
        return new Response(
          JSON.stringify({
            data: {
              Page: {
                media: [
                  {
                    id: 130298,
                    episodes: 20,
                    title: {
                      romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                      english: 'The Eminence in Shadow',
                      native: '陰の実力者になりたくて！',
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
      }

      if (body.query?.includes('characters(page: $page')) {
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                  english: 'The Eminence in Shadow',
                  native: '陰の実力者になりたくて！',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'SUPPORTING',
                      node: {
                        id: 123,
                        description:
                          '__Race:__ Human Alexia Midgar is the second princess of the Kingdom of Midgar.',
                        image: {
                          large: 'https://example.com/alexia.png',
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
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        );
      }
    }

    if (url === 'https://example.com/alexia.png') {
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
        episode: 5,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_000,
    });

    const result = await runtime.generateForCurrentMedia();
    const termBank = JSON.parse(readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8')) as Array<
      [string, string, string, string, number, Array<string | Record<string, unknown>>, number, string]
    >;
    const alexia = termBank.find(([term]) => term === 'アレクシア');

    assert.ok(alexia, 'expected compact native-name variant for character');
    const glossary = alexia[5];
    assert.equal(glossary.length, 1);

    const entry = glossary[0] as {
      type: string;
      content: unknown[];
    };
    assert.equal(entry.type, 'structured-content');
    assert.equal(Array.isArray(entry.content), true);

    const image = entry.content[0] as Record<string, unknown>;
    assert.equal(image.tag, 'img');
    assert.equal(image.path, 'img/m130298-c123.png');
    assert.equal(image.sizeUnits, 'em');

    const descriptionLine = entry.content[5];
    assert.equal(
      descriptionLine,
      'Race: Human Alexia Midgar is the second princess of the Kingdom of Midgar.',
    );

    const topLevelImageGlossaryEntry = glossary.find(
      (item) => typeof item === 'object' && item !== null && (item as { type?: string }).type === 'image',
    );
    assert.equal(topLevelImageGlossaryEntry, undefined);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia adds kana aliases for romanized names when native name is kanji', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
      };

      if (body.query?.includes('Page(perPage: 10)')) {
        return new Response(
          JSON.stringify({
            data: {
              Page: {
                media: [
                  {
                    id: 20594,
                    episodes: 10,
                    title: {
                      romaji: 'Kono Subarashii Sekai ni Shukufuku wo!',
                      english: 'KONOSUBA -God’s blessing on this wonderful world!',
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
      }

      if (body.query?.includes('characters(page: $page')) {
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  romaji: 'Kono Subarashii Sekai ni Shukufuku wo!',
                  english: 'KONOSUBA -God’s blessing on this wonderful world!',
                  native: 'この素晴らしい世界に祝福を！',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'MAIN',
                      node: {
                        id: 1,
                        description: 'The protagonist.',
                        image: null,
                        name: {
                          full: 'Satou Kazuma',
                          native: '佐藤和真',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        );
      }
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/konosuba-s02e05.mkv',
      getCurrentMediaTitle: () => 'Konosuba S02E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'Konosuba',
        episode: 5,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_000,
    });

    const result = await runtime.generateForCurrentMedia();
    const termBank = JSON.parse(readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8')) as Array<
      [string, string, string, string, number, Array<string | Record<string, unknown>>, number, string]
    >;

    const kazuma = termBank.find(([term]) => term === 'カズマ');
    assert.ok(kazuma, 'expected katakana alias for romanized name');
    assert.equal(kazuma[1], 'かずま');

    const fullName = termBank.find(([term]) => term === 'サトウカズマ');
    assert.ok(fullName, 'expected compact full-name katakana alias for romanized name');
    assert.equal(fullName[1], 'さとうかずま');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getOrCreateCurrentSnapshot persists and reuses normalized snapshot data', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  let searchQueryCount = 0;
  let characterQueryCount = 0;

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
      };

      if (body.query?.includes('Page(perPage: 10)')) {
        searchQueryCount += 1;
        return new Response(
          JSON.stringify({
            data: {
              Page: {
                media: [
                  {
                    id: 130298,
                    episodes: 20,
                    title: {
                      romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                      english: 'The Eminence in Shadow',
                      native: '陰の実力者になりたくて！',
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
      }

      if (body.query?.includes('characters(page: $page')) {
        characterQueryCount += 1;
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                  english: 'The Eminence in Shadow',
                  native: '陰の実力者になりたくて！',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'MAIN',
                      node: {
                        id: 321,
                        description: 'Alpha is the second-in-command of Shadow Garden.',
                        image: {
                          large: 'https://example.com/alpha.png',
                          medium: null,
                        },
                        name: {
                          full: 'Alpha',
                          native: 'アルファ',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        );
      }
    }

    if (url === 'https://example.com/alpha.png') {
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
        episode: 5,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_100,
    });

    const first = await runtime.getOrCreateCurrentSnapshot();
    const second = await runtime.getOrCreateCurrentSnapshot();

    assert.equal(first.fromCache, false);
    assert.equal(second.fromCache, true);
    assert.equal(searchQueryCount, 2);
    assert.equal(characterQueryCount, 1);
    assert.equal(
      fs.existsSync(path.join(userDataPath, 'character-dictionaries', 'cache.json')),
      false,
    );

    const snapshotPath = path.join(
      userDataPath,
      'character-dictionaries',
      'snapshots',
      'anilist-130298.json',
    );
    const snapshot = JSON.parse(fs.readFileSync(snapshotPath, 'utf8')) as {
      mediaId: number;
      entryCount: number;
      termEntries: Array<
        [string, string, string, string, number, Array<string | Record<string, unknown>>, number, string]
      >;
    };
    assert.equal(snapshot.mediaId, 130298);
    assert.equal(snapshot.entryCount > 0, true);
    const alpha = snapshot.termEntries.find(([term]) => term === 'アルファ');
    assert.ok(alpha);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getOrCreateCurrentSnapshot rebuilds snapshots written with an older format version', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  let searchQueryCount = 0;
  let characterQueryCount = 0;

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
      };

      if (body.query?.includes('Page(perPage: 10)')) {
        searchQueryCount += 1;
        return new Response(
          JSON.stringify({
            data: {
              Page: {
                media: [
                  {
                    id: 130298,
                    episodes: 20,
                    title: {
                      romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                      english: 'The Eminence in Shadow',
                      native: '陰の実力者になりたくて！',
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
      }

      if (body.query?.includes('characters(page: $page')) {
        characterQueryCount += 1;
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                  english: 'The Eminence in Shadow',
                  native: '陰の実力者になりたくて！',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'MAIN',
                      node: {
                        id: 321,
                        description: 'Alpha is the second-in-command of Shadow Garden.',
                        image: null,
                        name: {
                          full: 'Alpha',
                          native: 'アルファ',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        );
      }
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const snapshotsDir = path.join(userDataPath, 'character-dictionaries', 'snapshots');
    fs.mkdirSync(snapshotsDir, { recursive: true });
    fs.writeFileSync(
      path.join(snapshotsDir, 'anilist-130298.json'),
      JSON.stringify({
        formatVersion: 9,
        mediaId: 130298,
        mediaTitle: 'The Eminence in Shadow',
        entryCount: 1,
        updatedAt: 1_700_000_000_000,
        termEntries: [['stale', '', 'name side', '', 1, ['stale'], 0, '']],
        images: [],
      }),
      'utf8',
    );

    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        episode: 5,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_100,
    });

    const result = await runtime.getOrCreateCurrentSnapshot();

    assert.equal(result.fromCache, false);
    assert.equal(searchQueryCount, 1);
    assert.equal(characterQueryCount, 1);

    const snapshotPath = path.join(snapshotsDir, 'anilist-130298.json');
    const snapshot = JSON.parse(fs.readFileSync(snapshotPath, 'utf8')) as {
      formatVersion: number;
      termEntries: Array<
        [string, string, string, string, number, Array<string | Record<string, unknown>>, number, string]
      >;
    };
    assert.equal(snapshot.formatVersion > 9, true);
    assert.equal(snapshot.termEntries.some(([term]) => term === 'アルファ'), true);
    assert.equal(snapshot.termEntries.some(([term]) => term === 'stale'), false);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia logs progress while resolving and rebuilding snapshot data', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  const logs: string[] = [];

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
      };

      if (body.query?.includes('Page(perPage: 10)')) {
        return new Response(
          JSON.stringify({
            data: {
              Page: {
                media: [
                  {
                    id: 130298,
                    episodes: 20,
                    title: {
                      romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                      english: 'The Eminence in Shadow',
                      native: '陰の実力者になりたくて！',
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
      }

      if (body.query?.includes('characters(page: $page')) {
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                  english: 'The Eminence in Shadow',
                  native: '陰の実力者になりたくて！',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'MAIN',
                      node: {
                        id: 321,
                        description: 'Alpha is the second-in-command of Shadow Garden.',
                        image: {
                          large: 'https://example.com/alpha.png',
                          medium: null,
                        },
                        name: {
                          full: 'Alpha',
                          native: 'アルファ',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        );
      }
    }

    if (url === 'https://example.com/alpha.png') {
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
        episode: 5,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_100,
      sleep: async () => undefined,
      logInfo: (message) => {
        logs.push(message);
      },
    });

    await runtime.generateForCurrentMedia();

    assert.deepEqual(logs, [
      '[dictionary] resolving current anime for character dictionary generation',
      '[dictionary] current anime guess: The Eminence in Shadow (episode 5)',
      '[dictionary] AniList match: The Eminence in Shadow -> AniList 130298',
      '[dictionary] snapshot miss for AniList 130298, fetching characters',
      '[dictionary] downloaded AniList character page 1 for AniList 130298',
      '[dictionary] downloading 1 character images for AniList 130298',
      '[dictionary] stored snapshot for AniList 130298: 32 terms',
      '[dictionary] building ZIP for AniList 130298',
      '[dictionary] generated AniList 130298: 32 terms -> ' +
        path.join(userDataPath, 'character-dictionaries', 'anilist-130298.zip'),
    ]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('buildMergedDictionary combines stored snapshots into one stable dictionary', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  const current = { title: 'The Eminence in Shadow', episode: 5 };

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
        variables?: Record<string, unknown>;
      };

      if (body.query?.includes('Page(perPage: 10)')) {
        if (body.variables?.search === 'The Eminence in Shadow') {
          return new Response(
            JSON.stringify({
              data: {
                Page: {
                  media: [
                    {
                      id: 130298,
                      episodes: 20,
                      title: {
                        romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                        english: 'The Eminence in Shadow',
                        native: '陰の実力者になりたくて！',
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
        }

        return new Response(
          JSON.stringify({
            data: {
              Page: {
                media: [
                  {
                    id: 21,
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
      }

      if (body.query?.includes('characters(page: $page')) {
        const mediaId = Number(body.variables?.id);
        if (mediaId === 130298) {
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
                        role: 'MAIN',
                        node: {
                          id: 111,
                          description: 'Leader of Shadow Garden.',
                          image: {
                            large: 'https://example.com/alpha.png',
                            medium: null,
                          },
                          name: {
                            full: 'Alpha',
                            native: 'アルファ',
                          },
                        },
                      },
                    ],
                  },
                },
              },
            }),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          );
        }

        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  english: 'Frieren: Beyond Journey’s End',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'MAIN',
                      node: {
                        id: 222,
                        description: 'Elven mage.',
                        image: {
                          large: 'https://example.com/frieren.png',
                          medium: null,
                        },
                        name: {
                          full: 'Frieren',
                          native: 'フリーレン',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        );
      }
    }

    if (url === 'https://example.com/alpha.png' || url === 'https://example.com/frieren.png') {
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
      getCurrentMediaPath: () => '/tmp/current.mkv',
      getCurrentMediaTitle: () => current.title,
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: current.title,
        episode: current.episode,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_100,
    });

    await runtime.getOrCreateCurrentSnapshot();
    current.title = 'Frieren: Beyond Journey’s End';
    current.episode = 1;
    await runtime.getOrCreateCurrentSnapshot();

    const merged = await runtime.buildMergedDictionary([21, 130298]);
    const index = JSON.parse(readStoredZipEntry(merged.zipPath, 'index.json').toString('utf8')) as {
      title: string;
    };
    const termBank = JSON.parse(readStoredZipEntry(merged.zipPath, 'term_bank_1.json').toString('utf8')) as Array<
      [string, string, string, string, number, Array<string | Record<string, unknown>>, number, string]
    >;
    const frieren = termBank.find(([term]) => term === 'フリーレン');
    const alpha = termBank.find(([term]) => term === 'アルファ');

    assert.equal(index.title, 'SubMiner Character Dictionary');
    assert.equal(merged.entryCount >= 2, true);
    assert.ok(frieren);
    assert.ok(alpha);
    assert.equal((frieren[5][0] as { type?: string }).type, 'structured-content');
    assert.equal((alpha[5][0] as { type?: string }).type, 'structured-content');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia paces AniList requests and character image downloads', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  const sleepCalls: number[] = [];
  const imageRequests: string[] = [];

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
      };

      if (body.query?.includes('Page(perPage: 10)')) {
        return new Response(
          JSON.stringify({
            data: {
              Page: {
                media: [
                  {
                    id: 130298,
                    episodes: 20,
                    title: {
                      romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                      english: 'The Eminence in Shadow',
                      native: '陰の実力者になりたくて！',
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
      }

      if (body.query?.includes('characters(page: $page')) {
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  romaji: 'Kage no Jitsuryokusha ni Naritakute!',
                  english: 'The Eminence in Shadow',
                  native: '陰の実力者になりたくて！',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'MAIN',
                      node: {
                        id: 111,
                        description: 'First character.',
                        image: {
                          large: 'https://example.com/alpha.png',
                          medium: null,
                        },
                        name: {
                          full: 'Alpha',
                          native: 'アルファ',
                        },
                      },
                    },
                    {
                      role: 'SUPPORTING',
                      node: {
                        id: 222,
                        description: 'Second character.',
                        image: {
                          large: 'https://example.com/beta.png',
                          medium: null,
                        },
                        name: {
                          full: 'Beta',
                          native: 'ベータ',
                        },
                      },
                    },
                  ],
                },
              },
            },
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        );
      }
    }

    if (url === 'https://example.com/alpha.png') {
      imageRequests.push(url);
      return new Response('missing', {
        status: 404,
        headers: { 'content-type': 'text/plain' },
      });
    }

    if (url === 'https://example.com/beta.png') {
      imageRequests.push(url);
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
        episode: 5,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_000,
      sleep: async (ms) => {
        sleepCalls.push(ms);
      },
    });

    await runtime.generateForCurrentMedia();

    assert.deepEqual(sleepCalls, [2000, 250]);
    assert.deepEqual(imageRequests, ['https://example.com/alpha.png', 'https://example.com/beta.png']);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
