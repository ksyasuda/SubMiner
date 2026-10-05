import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { TestContext } from 'node:test';

import { createCharacterDictionaryRuntimeService } from '../character-dictionary-runtime';
import { ANILIST_GRAPHQL_URL } from './constants';
import type { CharacterDictionaryRuntimeDeps, CharacterDictionaryTermEntry } from './types';

// Shared fixtures for the character dictionary runtime tests: a fake AniList GraphQL + image
// server, a runtime factory with Eminence in Shadow defaults, and readers for the built ZIP.

export const PNG_1X1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nmX8AAAAASUVORK5CYII=',
  'base64',
);

export function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-dictionary-'));
}

type FetchHandler = (url: string, init?: RequestInit) => Promise<Response> | Response;

/** Replaces `globalThis.fetch` for the duration of the test. */
export function installFetch(t: TestContext, handler: FetchHandler): void {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    return handler(url, init);
  }) as typeof globalThis.fetch;
  t.after(() => {
    globalThis.fetch = originalFetch;
  });
}

export function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

/** AniList `characters.edges[]` item, as returned by the GraphQL API. */
export type FakeAniListCharacterEdge = {
  role: string;
  voiceActors?: Array<{
    id: number;
    name: { full?: string; native?: string };
    image?: { large?: string | null; medium?: string | null };
  }>;
  node: {
    id: number;
    description?: string;
    image?: { large?: string | null; medium?: string | null } | null;
    name: {
      full?: string;
      native?: string;
      first?: string | null;
      last?: string | null;
      alternative?: string[];
    };
    gender?: string;
    age?: string;
    dateOfBirth?: { month: number; day: number };
    bloodType?: string;
  };
};

export type FakeAniListMedia = {
  id: number;
  episodes?: number;
  title: { romaji?: string; english?: string; native?: string };
  /** Search text that resolves to this media. Unmatched searches resolve to the first media. */
  search?: string;
  characters?: FakeAniListCharacterEdge[];
};

export type FakeAniList = {
  /** Every fetched URL, in request order. */
  requests: string[];
  searchCount: number;
  characterPageCount: number;
  imageRequests: string[];
};

/**
 * Serves AniList media search and character page queries from `media`, and image downloads for
 * `images` (PNG) and `missingImages` (404). Any other request throws.
 */
export function installFakeAniList(
  t: TestContext,
  options: { media?: FakeAniListMedia[]; images?: string[]; missingImages?: string[] } = {},
): FakeAniList {
  const media = options.media ?? [];
  const images = new Set(options.images ?? []);
  const missingImages = new Set(options.missingImages ?? []);
  const fake: FakeAniList = {
    requests: [],
    searchCount: 0,
    characterPageCount: 0,
    imageRequests: [],
  };

  installFetch(t, (url, init) => {
    fake.requests.push(url);
    if (url === ANILIST_GRAPHQL_URL) {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        query?: string;
        variables?: { search?: string; id?: number };
      };

      if (body.query?.includes('Page(perPage: 10)') && media.length > 0) {
        fake.searchCount += 1;
        const match = media.find((entry) => entry.search === body.variables?.search) ?? media[0]!;
        return jsonResponse({
          data: {
            Page: { media: [{ id: match.id, episodes: match.episodes ?? 12, title: match.title }] },
          },
        });
      }

      const pageMedia = media.find((entry) => entry.id === body.variables?.id);
      if (body.query?.includes('characters(page: $page') && pageMedia) {
        fake.characterPageCount += 1;
        return jsonResponse({
          data: {
            Media: {
              title: pageMedia.title,
              characters: {
                pageInfo: { hasNextPage: false },
                edges: pageMedia.characters ?? [],
              },
            },
          },
        });
      }
    }

    if (images.has(url) || missingImages.has(url)) {
      fake.imageRequests.push(url);
      return images.has(url)
        ? new Response(PNG_1X1, { status: 200, headers: { 'content-type': 'image/png' } })
        : new Response('missing', { status: 404, headers: { 'content-type': 'text/plain' } });
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  });

  return fake;
}

export const EMINENCE_MEDIA_ID = 130298;
export const EMINENCE_TITLE = {
  romaji: 'Kage no Jitsuryokusha ni Naritakute!',
  english: 'The Eminence in Shadow',
  native: '陰の実力者になりたくて！',
};

export type CharacterDictionaryRuntime = ReturnType<typeof createCharacterDictionaryRuntimeService>;

/** Creates a runtime watching an Eminence in Shadow episode. AniList pacing sleeps are skipped. */
export function createRuntime(overrides: Partial<CharacterDictionaryRuntimeDeps> = {}): {
  runtime: CharacterDictionaryRuntime;
  userDataPath: string;
  outputDir: string;
} {
  const userDataPath = overrides.userDataPath ?? makeTempDir();
  const runtime = createCharacterDictionaryRuntimeService({
    getCurrentMediaPath: () => '/tmp/eminence-s01e05.mkv',
    getCurrentMediaTitle: () => 'The Eminence in Shadow - S01E05',
    resolveMediaPathForJimaku: (mediaPath) => mediaPath,
    guessAnilistMediaInfo: async () => ({
      title: 'The Eminence in Shadow',
      season: null,
      episode: 5,
      source: 'fallback',
    }),
    now: () => 1_700_000_000_000,
    sleep: async () => undefined,
    ...overrides,
    userDataPath,
  });
  return { runtime, userDataPath, outputDir: path.join(userDataPath, 'character-dictionaries') };
}

const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;

export function readStoredZipEntry(zipPath: string, entryName: string): Buffer {
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

export function readTermBank(zipPath: string): CharacterDictionaryTermEntry[] {
  return JSON.parse(
    readStoredZipEntry(zipPath, 'term_bank_1.json').toString('utf8'),
  ) as CharacterDictionaryTermEntry[];
}

type StructuredNode = { tag?: string; open?: boolean; content?: unknown };

/** Top-level structured-content children of a term entry's glossary. */
export function getGlossaryChildren(entry: CharacterDictionaryTermEntry): StructuredNode[] {
  const glossary = entry[5][0] as { content: { content: StructuredNode[] } };
  return glossary.content.content;
}

/** Finds a collapsible `details` section by its summary title. */
export function findSection(
  children: StructuredNode[],
  title: string,
): (StructuredNode & { content: StructuredNode[] }) | undefined {
  return children.find(
    (child): child is StructuredNode & { content: StructuredNode[] } =>
      child.tag === 'details' &&
      Array.isArray(child.content) &&
      (child.content[0] as StructuredNode | undefined)?.content === title,
  );
}
