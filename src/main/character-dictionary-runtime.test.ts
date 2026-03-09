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
                          '__Race:__ Human\nAlexia Midgar is the second princess of the Kingdom of Midgar.',
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
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;
    const alexia = termBank.find(([term]) => term === 'アレクシア');

    assert.ok(alexia, 'expected compact native-name variant for character');
    const glossary = alexia[5];
    assert.equal(glossary.length, 1);

    const entry = glossary[0] as {
      type: string;
      content: { tag: string; content: Array<Record<string, unknown>> };
    };
    assert.equal(entry.type, 'structured-content');

    const wrapper = entry.content;
    assert.equal(wrapper.tag, 'div');
    const children = wrapper.content;

    const nameDiv = children[0] as { tag: string; content: string };
    assert.equal(nameDiv.tag, 'div');
    assert.equal(nameDiv.content, 'アレクシア・ミドガル');

    const secondaryNameDiv = children[1] as { tag: string; content: string };
    assert.equal(secondaryNameDiv.tag, 'div');
    assert.equal(secondaryNameDiv.content, 'Alexia Midgar');

    const imageWrap = children[2] as { tag: string; content: Record<string, unknown> };
    assert.equal(imageWrap.tag, 'div');
    const image = imageWrap.content as Record<string, unknown>;
    assert.equal(image.tag, 'img');
    assert.equal(image.path, 'img/m130298-c123.png');
    assert.equal(image.sizeUnits, 'em');

    const sourceDiv = children[3] as { tag: string; content: string };
    assert.equal(sourceDiv.tag, 'div');
    assert.ok(sourceDiv.content.includes('The Eminence in Shadow'));

    const roleBadgeDiv = children[4] as { tag: string; content: Record<string, unknown> };
    assert.equal(roleBadgeDiv.tag, 'div');
    const badge = roleBadgeDiv.content as { tag: string; content: string };
    assert.equal(badge.tag, 'span');
    assert.equal(badge.content, 'Main Character');

    const descSection = children.find(
      (c) =>
        (c as { tag?: string }).tag === 'details' &&
        Array.isArray((c as { content?: unknown[] }).content) &&
        (c as { content: Array<{ content?: string }> }).content[0]?.content === 'Description',
    ) as { tag: string; open?: boolean; content: Array<Record<string, unknown>> } | undefined;
    assert.ok(descSection, 'expected Description collapsible section');
    assert.equal(descSection.open, false);
    const descBody = descSection.content[1] as { content: string };
    assert.ok(
      descBody.content.includes('Alexia Midgar is the second princess of the Kingdom of Midgar.'),
    );

    const infoSection = children.find(
      (c) =>
        (c as { tag?: string }).tag === 'details' &&
        Array.isArray((c as { content?: unknown[] }).content) &&
        (c as { content: Array<{ content?: string }> }).content[0]?.content ===
          'Character Information',
    ) as { tag: string; open?: boolean; content: Array<Record<string, unknown>> } | undefined;
    assert.ok(
      infoSection,
      'expected Character Information collapsible section with parsed __Race:__ field',
    );
    assert.equal(infoSection.open, false);

    const topLevelImageGlossaryEntry = glossary.find(
      (item) =>
        typeof item === 'object' && item !== null && (item as { type?: string }).type === 'image',
    );
    assert.equal(topLevelImageGlossaryEntry, undefined);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia applies configured open states to character dictionary sections', async () => {
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
                    id: 130298,
                    episodes: 20,
                    title: {
                      romaji: 'The Eminence in Shadow',
                      english: 'The Eminence in Shadow',
                      native: '陰の実力者になりたくて！',
                    },
                  },
                ],
              },
            },
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }

      if (body.query?.includes('characters(page: $page')) {
        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: {
                  romaji: 'The Eminence in Shadow',
                  english: 'The Eminence in Shadow',
                  native: '陰の実力者になりたくて！',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'SUPPORTING',
                      voiceActors: [
                        {
                          id: 456,
                          name: {
                            full: 'Rina Hidaka',
                            native: '日高里菜',
                          },
                          image: {
                            medium: 'https://cdn.example.com/va-456.jpg',
                          },
                        },
                      ],
                      node: {
                        id: 123,
                        description:
                          'Alexia Midgar is the second princess of the Kingdom of Midgar.\n\n__Race:__ Human',
                        image: {
                          large: 'https://cdn.example.com/character-123.png',
                          medium: 'https://cdn.example.com/character-123-small.png',
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
      return new Response(Buffer.from([0x89, 0x50, 0x4e, 0x47]), {
        status: 200,
        headers: { 'content-type': 'image/png' },
      });
    }

    if (url === 'https://cdn.example.com/va-456.jpg') {
      return new Response(Buffer.from([0xff, 0xd8, 0xff, 0xd9]), {
        status: 200,
        headers: { 'content-type': 'image/jpeg' },
      });
    }

    throw new Error(`Unexpected fetch: ${url}`);
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
      getCollapsibleSectionOpenState: (section) =>
        section === 'description' || section === 'voicedBy',
      now: () => 1_700_000_000_000,
    });

    const result = await runtime.generateForCurrentMedia();
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;
    const alexia = termBank.find(([term]) => term === 'アレクシア');
    assert.ok(alexia);

    const glossary = alexia[5];
    const entry = glossary[0] as {
      type: string;
      content: { tag: string; content: Array<Record<string, unknown>> };
    };
    const children = entry.content.content;

    const getSection = (title: string) =>
      children.find(
        (c) =>
          (c as { tag?: string }).tag === 'details' &&
          Array.isArray((c as { content?: unknown[] }).content) &&
          (c as { content: Array<{ content?: string }> }).content[0]?.content === title,
      ) as { open?: boolean } | undefined;

    assert.equal(getSection('Description')?.open, true);
    assert.equal(getSection('Character Information')?.open, false);
    assert.equal(getSection('Voiced by')?.open, true);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia reapplies collapsible open states when using cached snapshot data', async () => {
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
                    id: 130298,
                    episodes: 20,
                    title: {
                      romaji: 'The Eminence in Shadow',
                      english: 'The Eminence in Shadow',
                      native: '陰の実力者になりたくて！',
                    },
                  },
                ],
              },
            },
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }

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
                      voiceActors: [
                        {
                          id: 456,
                          name: {
                            full: 'Rina Hidaka',
                            native: '日高里菜',
                          },
                          image: {
                            medium: 'https://cdn.example.com/va-456.jpg',
                          },
                        },
                      ],
                      node: {
                        id: 123,
                        description:
                          'Alexia Midgar is the second princess of the Kingdom of Midgar.\n\n__Race:__ Human',
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

    if (url === 'https://cdn.example.com/va-456.jpg') {
      return new Response(PNG_1X1, {
        status: 200,
        headers: { 'content-type': 'image/png' },
      });
    }

    throw new Error(`Unexpected fetch: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const runtimeOpen = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        episode: 5,
        source: 'fallback',
      }),
      getCollapsibleSectionOpenState: () => true,
      now: () => 1_700_000_000_000,
    });
    await runtimeOpen.generateForCurrentMedia();

    const runtimeClosed = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
      getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'The Eminence in Shadow',
        episode: 5,
        source: 'fallback',
      }),
      getCollapsibleSectionOpenState: () => false,
      now: () => 1_700_000_000_500,
    });
    const result = await runtimeClosed.generateForCurrentMedia();

    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;
    const alexia = termBank.find(([term]) => term === 'アレクシア');
    assert.ok(alexia);

    const children = (
      alexia[5][0] as {
        content: { content: Array<Record<string, unknown>> };
      }
    ).content.content;
    const sections = children.filter(
      (item) => (item as { tag?: string }).tag === 'details',
    ) as Array<{
      open?: boolean;
    }>;
    assert.ok(sections.length >= 2);
    assert.ok(sections.every((section) => section.open === false));
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
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
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

test('generateForCurrentMedia indexes kanji family and given names using AniList first and last hints', async () => {
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
                    id: 37450,
                    episodes: 13,
                    title: {
                      romaji: 'Seishun Buta Yarou wa Bunny Girl Senpai no Yume wo Minai',
                      english: 'Rascal Does Not Dream of Bunny Girl Senpai',
                      native: '青春ブタ野郎はバニーガール先輩の夢を見ない',
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
                  romaji: 'Seishun Buta Yarou wa Bunny Girl Senpai no Yume wo Minai',
                  english: 'Rascal Does Not Dream of Bunny Girl Senpai',
                  native: '青春ブタ野郎はバニーガール先輩の夢を見ない',
                },
                characters: {
                  pageInfo: { hasNextPage: false },
                  edges: [
                    {
                      role: 'SUPPORTING',
                      node: {
                        id: 77,
                        description: 'Classmate.',
                        image: null,
                        name: {
                          first: 'Yuuma',
                          full: 'Yuuma Kunimi',
                          last: 'Kunimi',
                          native: '国見佑真',
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
      getCurrentMediaPath: () => '/tmp/bunny-girl-senpai-s01e01.mkv',
      getCurrentMediaTitle: () => 'Rascal Does Not Dream of Bunny Girl Senpai - S01E01',
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: 'Rascal Does Not Dream of Bunny Girl Senpai',
        episode: 1,
        source: 'fallback',
      }),
      now: () => 1_700_000_000_000,
    });

    const result = await runtime.generateForCurrentMedia();
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;

    const familyName = termBank.find(([term]) => term === '国見');
    assert.ok(familyName, 'expected kanji family-name term from AniList hints');
    assert.equal(familyName[1], 'くにみ');

    const givenName = termBank.find(([term]) => term === '佑真');
    assert.ok(givenName, 'expected kanji given-name term from AniList hints');
    assert.equal(givenName[1], 'ゆうま');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia indexes AniList alternative character names for alias lookups', async () => {
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
                        description: 'Leader of Shadow Garden.',
                        image: null,
                        name: {
                          full: 'Cid Kagenou',
                          native: 'シド・カゲノー',
                          alternative: ['Shadow', 'Minoru Kagenou'],
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
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;

    const shadowKana = termBank.find(([term]) => term === 'シャドウ');
    assert.ok(shadowKana, 'expected katakana alias from AniList alternative name');
    assert.equal(shadowKana[1], 'しゃどう');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia skips AniList characters without a native name when other valid characters exist', async () => {
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
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }

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
                      role: 'MAIN',
                      node: {
                        id: 111,
                        description: 'Valid native name.',
                        image: null,
                        name: {
                          full: 'Alpha',
                          native: 'アルファ',
                          first: 'Alpha',
                          last: null,
                        },
                      },
                    },
                    {
                      role: 'SUPPORTING',
                      node: {
                        id: 222,
                        description: 'Missing native name.',
                        image: null,
                        name: {
                          full: 'John Smith',
                          native: '',
                          first: 'John',
                          last: 'Smith',
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
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;

    assert.ok(termBank.find(([term]) => term === 'アルファ'));
    assert.equal(
      termBank.some(([term]) => term === 'John Smith'),
      false,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia uses AniList first and last name hints to build kanji readings', async () => {
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
                          first: '和真',
                          last: '佐藤',
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
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;

    assert.equal(termBank.find(([term]) => term === '佐藤和真')?.[1], 'さとうかずま');
    assert.equal(termBank.find(([term]) => term === '佐藤')?.[1], 'さとう');
    assert.equal(termBank.find(([term]) => term === '和真')?.[1], 'かずま');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia includes AniList gender age birthday and blood type in character information', async () => {
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
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }

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
                        description: 'Second princess of Midgar.',
                        image: null,
                        gender: 'Female',
                        age: '15',
                        dateOfBirth: {
                          month: 9,
                          day: 1,
                        },
                        bloodType: 'A',
                        name: {
                          full: 'Alexia Midgar',
                          native: 'アレクシア・ミドガル',
                          first: 'Alexia',
                          last: 'Midgar',
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
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;
    const alexia = termBank.find(([term]) => term === 'アレクシア');
    assert.ok(alexia);

    const children = (
      alexia[5][0] as {
        content: { content: Array<Record<string, unknown>> };
      }
    ).content.content;
    const infoSection = children.find(
      (c) =>
        (c as { tag?: string }).tag === 'details' &&
        Array.isArray((c as { content?: unknown[] }).content) &&
        (c as { content: Array<{ content?: string }> }).content[0]?.content ===
          'Character Information',
    ) as { content: Array<Record<string, unknown>> } | undefined;
    assert.ok(infoSection);
    const body = infoSection.content[1] as { content: Array<{ content?: string }> };
    const flattened = JSON.stringify(body.content);

    assert.match(flattened, /Female|♂ Male|♀ Female/);
    assert.match(flattened, /15 years/);
    assert.match(flattened, /Blood Type A/);
    assert.match(flattened, /Birthday: September 1/);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia preserves duplicate surface forms across different characters', async () => {
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
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }

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
                      role: 'MAIN',
                      node: {
                        id: 111,
                        description: 'First Alpha.',
                        image: null,
                        name: {
                          full: 'Alpha One',
                          native: 'アルファ',
                          first: 'Alpha',
                          last: 'One',
                        },
                      },
                    },
                    {
                      role: 'MAIN',
                      node: {
                        id: 222,
                        description: 'Second Alpha.',
                        image: null,
                        name: {
                          full: 'Alpha Two',
                          native: 'アルファ',
                          first: 'Alpha',
                          last: 'Two',
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
    const termBank = JSON.parse(
      readStoredZipEntry(result.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;

    const alphaEntries = termBank.filter(([term]) => term === 'アルファ');
    assert.equal(alphaEntries.length, 2);
    const glossaries = alphaEntries.map((entry) =>
      JSON.stringify(
        (
          entry[5][0] as {
            content: { content: Array<Record<string, unknown>> };
          }
        ).content.content,
      ),
    );
    assert.ok(glossaries.some((value) => value.includes('First Alpha.')));
    assert.ok(glossaries.some((value) => value.includes('Second Alpha.')));
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
        [
          string,
          string,
          string,
          string,
          number,
          Array<string | Record<string, unknown>>,
          number,
          string,
        ]
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
        [
          string,
          string,
          string,
          string,
          number,
          Array<string | Record<string, unknown>>,
          number,
          string,
        ]
      >;
    };
    assert.equal(snapshot.formatVersion > 9, true);
    assert.equal(
      snapshot.termEntries.some(([term]) => term === 'アルファ'),
      true,
    );
    assert.equal(
      snapshot.termEntries.some(([term]) => term === 'stale'),
      false,
    );
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
      '[dictionary] downloading 1 images for AniList 130298',
      '[dictionary] stored snapshot for AniList 130298: 32 terms',
      '[dictionary] building ZIP for AniList 130298',
      '[dictionary] generated AniList 130298: 32 terms -> ' +
        path.join(userDataPath, 'character-dictionaries', 'anilist-130298.zip'),
    ]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('generateForCurrentMedia downloads shared voice actor images once per AniList person id', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  const fetchedImageUrls: string[] = [];

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
                      voiceActors: [
                        {
                          id: 9001,
                          name: {
                            full: 'Kana Hanazawa',
                            native: '花澤香菜',
                          },
                          image: {
                            large: null,
                            medium: 'https://example.com/kana.png',
                          },
                        },
                      ],
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
                    {
                      role: 'SUPPORTING',
                      voiceActors: [
                        {
                          id: 9001,
                          name: {
                            full: 'Kana Hanazawa',
                            native: '花澤香菜',
                          },
                          image: {
                            large: null,
                            medium: 'https://example.com/kana.png',
                          },
                        },
                      ],
                      node: {
                        id: 654,
                        description: 'Beta documents Shadow Garden operations.',
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

    if (
      url === 'https://example.com/alpha.png' ||
      url === 'https://example.com/beta.png' ||
      url === 'https://example.com/kana.png'
    ) {
      fetchedImageUrls.push(url);
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
    });

    await runtime.generateForCurrentMedia();

    assert.deepEqual(fetchedImageUrls, [
      'https://example.com/alpha.png',
      'https://example.com/kana.png',
      'https://example.com/beta.png',
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
    const termBank = JSON.parse(
      readStoredZipEntry(merged.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
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

test('buildMergedDictionary rebuilds snapshots written with an older format version', async () => {
  const userDataPath = makeTempDir();
  const originalFetch = globalThis.fetch;
  let characterQueryCount = 0;

  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (url === GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
        variables?: Record<string, unknown>;
      };

      if (body.query?.includes('characters(page: $page')) {
        characterQueryCount += 1;
        assert.equal(body.variables?.id, 130298);
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
                        image: null,
                        name: {
                          full: 'Cid Kagenou',
                          native: 'シド・カゲノー',
                          alternative: ['Shadow'],
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
    const snapshotsDir = path.join(userDataPath, 'character-dictionaries', 'snapshots');
    fs.mkdirSync(snapshotsDir, { recursive: true });
    fs.writeFileSync(
      path.join(snapshotsDir, 'anilist-130298.json'),
      JSON.stringify({
        formatVersion: 12,
        mediaId: 130298,
        mediaTitle: 'The Eminence in Shadow',
        entryCount: 1,
        updatedAt: 1_700_000_000_000,
        termEntries: [['stale', '', 'name main', '', 100, ['stale'], 0, '']],
        images: [],
      }),
      'utf8',
    );

    const runtime = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => null,
      getCurrentMediaTitle: () => null,
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => null,
      now: () => 1_700_000_000_100,
    });

    const merged = await runtime.buildMergedDictionary([130298]);
    const termBank = JSON.parse(
      readStoredZipEntry(merged.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;

    assert.equal(characterQueryCount, 1);
    assert.ok(termBank.find(([term]) => term === 'シャドウ'));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('buildMergedDictionary reapplies collapsible open states from current config', async () => {
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
                        english: 'The Eminence in Shadow',
                      },
                    },
                  ],
                },
              },
            }),
            { status: 200, headers: { 'content-type': 'application/json' } },
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
                      english: 'Frieren: Beyond Journey’s End',
                    },
                  },
                ],
              },
            },
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        );
      }

      if (body.query?.includes('characters(page: $page')) {
        const mediaId = Number(body.variables?.id);
        if (mediaId === 130298) {
          return new Response(
            JSON.stringify({
              data: {
                Media: {
                  title: { english: 'The Eminence in Shadow' },
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
            { status: 200, headers: { 'content-type': 'application/json' } },
          );
        }

        return new Response(
          JSON.stringify({
            data: {
              Media: {
                title: { english: 'Frieren: Beyond Journey’s End' },
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
          { status: 200, headers: { 'content-type': 'application/json' } },
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
    const runtimeOpen = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/current.mkv',
      getCurrentMediaTitle: () => current.title,
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: current.title,
        episode: current.episode,
        source: 'fallback',
      }),
      getCollapsibleSectionOpenState: () => true,
      now: () => 1_700_000_000_100,
    });

    await runtimeOpen.getOrCreateCurrentSnapshot();
    current.title = 'Frieren: Beyond Journey’s End';
    current.episode = 1;
    await runtimeOpen.getOrCreateCurrentSnapshot();

    const runtimeClosed = createCharacterDictionaryRuntimeService({
      userDataPath,
      getCurrentMediaPath: () => '/tmp/current.mkv',
      getCurrentMediaTitle: () => current.title,
      resolveMediaPathForJimaku: (mediaPath) => mediaPath,
      guessAnilistMediaInfo: async () => ({
        title: current.title,
        episode: current.episode,
        source: 'fallback',
      }),
      getCollapsibleSectionOpenState: () => false,
      now: () => 1_700_000_000_200,
    });

    const merged = await runtimeClosed.buildMergedDictionary([21, 130298]);
    const termBank = JSON.parse(
      readStoredZipEntry(merged.zipPath, 'term_bank_1.json').toString('utf8'),
    ) as Array<
      [
        string,
        string,
        string,
        string,
        number,
        Array<string | Record<string, unknown>>,
        number,
        string,
      ]
    >;
    const alpha = termBank.find(([term]) => term === 'アルファ');
    assert.ok(alpha);
    const children = (
      alpha[5][0] as {
        content: { content: Array<Record<string, unknown>> };
      }
    ).content.content;
    const sections = children.filter(
      (item) => (item as { tag?: string }).tag === 'details',
    ) as Array<{
      open?: boolean;
    }>;
    assert.ok(sections.length >= 1);
    assert.ok(sections.every((section) => section.open === false));
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
    assert.deepEqual(imageRequests, [
      'https://example.com/alpha.png',
      'https://example.com/beta.png',
    ]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
