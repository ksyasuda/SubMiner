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
    assert.equal(image.path, 'img/c123.png');
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

test('generateForCurrentMedia regenerates dictionary when cached format version is stale', async () => {
  const userDataPath = makeTempDir();
  const dictionariesDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(dictionariesDir, { recursive: true });

  const staleZipPath = path.join(dictionariesDir, 'anilist-130298.zip');
  fs.writeFileSync(staleZipPath, Buffer.from('not-a-real-zip'));
  fs.writeFileSync(
    path.join(dictionariesDir, 'cache.json'),
    JSON.stringify(
      {
        anilistById: {
          '130298': {
            mediaId: 130298,
            mediaTitle: 'The Eminence in Shadow',
            entryCount: 1,
            zipPath: staleZipPath,
            updatedAt: 1_700_000_000_000,
            formatVersion: 6,
            dictionaryTitle: 'SubMiner Character Dictionary (AniList 130298)',
            revision: 'stale-revision',
          },
        },
      },
      null,
      2,
    ),
    'utf8',
  );

  const originalFetch = globalThis.fetch;
  let characterQueryCount = 0;

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

    const result = await runtime.generateForCurrentMedia(undefined, {
      refreshTtlMs: 60 * 60 * 1000,
    });
    assert.equal(result.fromCache, false);
    assert.equal(characterQueryCount, 1);

    const termBank = JSON.parse(readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8')) as Array<
      [string, string, string, string, number, Array<string | Record<string, unknown>>, number, string]
    >;
    const alpha = termBank.find(([term]) => term === 'アルファ');
    assert.ok(alpha);
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
