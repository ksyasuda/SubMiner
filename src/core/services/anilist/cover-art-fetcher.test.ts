import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { createCoverArtFetcher, stripFilenameTags } from './cover-art-fetcher.js';
import { Database } from '../immersion-tracker/sqlite.js';
import { ensureSchema, getOrCreateVideoRecord } from '../immersion-tracker/storage.js';
import { getCoverArt } from '../immersion-tracker/query-library.js';
import { upsertCoverArt } from '../immersion-tracker/query-maintenance.js';
import { SOURCE_TYPE_LOCAL } from '../immersion-tracker/types.js';

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-cover-art-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
}

test('stripFilenameTags normalizes common media-title formats', () => {
  assert.equal(
    stripFilenameTags('[Jellyfin/direct] The Eminence in Shadow S01E05 I Am...'),
    'The Eminence in Shadow',
  );
  assert.equal(
    stripFilenameTags(
      '[Foxtrot] Kono Subarashii Sekai ni Shukufuku wo! S2 - 05: Servitude for this Masked Knight!',
    ),
    'Kono Subarashii Sekai ni Shukufuku wo!',
  );
  assert.equal(
    stripFilenameTags('Kono Subarashii Sekai ni Shukufuku wo! E03: A Panty Treasure'),
    'Kono Subarashii Sekai ni Shukufuku wo!',
  );
  assert.equal(
    stripFilenameTags(
      'Little Witch Academia (2017) - S01E05 - 005 - Pact of the Dragon [Bluray-1080p][10bit][h265][FLAC 2.0][JA]-FumeiRaws.mkv',
    ),
    'Little Witch Academia',
  );
});

test('fetchIfMissing backfills a missing blob from an existing cover URL', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  const videoId = getOrCreateVideoRecord(db, 'local:/tmp/cover-fetcher-test.mkv', {
    canonicalTitle: 'Cover Fetcher Test',
    sourcePath: '/tmp/cover-fetcher-test.mkv',
    sourceUrl: null,
    sourceType: SOURCE_TYPE_LOCAL,
  });
  upsertCoverArt(db, videoId, {
    anilistId: 7,
    coverUrl: 'https://images.test/cover.jpg',
    coverBlob: null,
    titleRomaji: 'Test Title',
    titleEnglish: 'Test Title',
    episodesTotal: 12,
  });

  const fetchCalls: string[] = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    fetchCalls.push(url);
    assert.equal(url, 'https://images.test/cover.jpg');
    return new Response(new Uint8Array([1, 2, 3, 4]), {
      status: 200,
      headers: { 'Content-Type': 'image/jpeg' },
    });
  }) as typeof fetch;

  try {
    const fetcher = createCoverArtFetcher(
      {
        acquire: async () => {},
        recordResponse: () => {},
      },
      console,
    );

    const fetched = await fetcher.fetchIfMissing(
      db,
      videoId,
      '[Jellyfin] Little Witch Academia S02E05 - 025 - Pact of the Dragon (2020) [1080p].mkv',
    );
    const stored = getCoverArt(db, videoId);

    assert.equal(fetched, true);
    assert.equal(fetchCalls.length, 1);
    assert.equal(stored?.coverBlob?.length, 4);
    assert.equal(stored?.titleEnglish, 'Test Title');
  } finally {
    globalThis.fetch = originalFetch;
    db.close();
    cleanupDbPath(dbPath);
  }
});

function createJsonResponse(payload: unknown): Response {
  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

test('fetchIfMissing uses guessit primary title and season when available', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  const videoId = getOrCreateVideoRecord(db, 'local:/tmp/cover-fetcher-season-test.mkv', {
    canonicalTitle:
      '[Jellyfin] Little Witch Academia S02E05 - 025 - Pact of the Dragon (2020) [1080p].mkv',
    sourcePath: '/tmp/cover-fetcher-season-test.mkv',
    sourceUrl: null,
    sourceType: SOURCE_TYPE_LOCAL,
  });

  const searchCalls: Array<{ search: string }> = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
    const raw = (init?.body as string | undefined) ?? '';
    const payload = JSON.parse(raw) as { variables: { search: string } };
    const search = payload.variables.search;
    searchCalls.push({ search });

    if (search.includes('Season 2')) {
      return Promise.resolve(createJsonResponse({ data: { Page: { media: [] } } }));
    }

    return Promise.resolve(
      createJsonResponse({
        data: {
          Page: {
            media: [
              {
                id: 19,
                episodes: 24,
                coverImage: { large: 'https://images.test/cover.jpg', medium: null },
                title: {
                  romaji: 'Little Witch Academia',
                  english: 'Little Witch Academia',
                  native: null,
                },
              },
            ],
          },
        },
      }),
    );
  }) as typeof fetch;

  try {
    const fetcher = createCoverArtFetcher(
      {
        acquire: async () => {},
        recordResponse: () => {},
      },
      console,
      {
        runGuessit: async () =>
          JSON.stringify({ title: 'Little Witch Academia', season: 2, episode: 5 }),
      },
    );

    const fetched = await fetcher.fetchIfMissing(db, videoId, 'School Vlog S01E01');
    const stored = getCoverArt(db, videoId);

    assert.equal(fetched, true);
    assert.equal(searchCalls.length, 2);
    assert.equal(searchCalls[0]!.search, 'Little Witch Academia Season 2');
    assert.equal(stored?.anilistId, 19);
  } finally {
    globalThis.fetch = originalFetch;
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('fetchIfMissing falls back to internal parser when guessit throws', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  const videoId = getOrCreateVideoRecord(db, 'local:/tmp/cover-fetcher-fallback-test.mkv', {
    canonicalTitle: 'School Vlog S01E01',
    sourcePath: '/tmp/cover-fetcher-fallback-test.mkv',
    sourceUrl: null,
    sourceType: SOURCE_TYPE_LOCAL,
  });

  let requestCount = 0;
  const originalFetch = globalThis.fetch;
  globalThis.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
    requestCount += 1;
    const raw = (init?.body as string | undefined) ?? '';
    const payload = JSON.parse(raw) as { variables: { search: string } };
    assert.equal(payload.variables.search, 'School Vlog');

    return Promise.resolve(
      createJsonResponse({
        data: {
          Page: {
            media: [
              {
                id: 21,
                episodes: 12,
                coverImage: { large: 'https://images.test/fallback-cover.jpg', medium: null },
                title: { romaji: 'School Vlog', english: 'School Vlog', native: null },
              },
            ],
          },
        },
      }),
    );
  }) as typeof fetch;

  try {
    const fetcher = createCoverArtFetcher(
      {
        acquire: async () => {},
        recordResponse: () => {},
      },
      console,
      {
        runGuessit: async () => {
          throw new Error('guessit unavailable');
        },
      },
    );

    const fetched = await fetcher.fetchIfMissing(db, videoId, 'Ignored Title');
    const stored = getCoverArt(db, videoId);

    assert.equal(fetched, true);
    assert.equal(requestCount, 2);
    assert.equal(stored?.anilistId, 21);
  } finally {
    globalThis.fetch = originalFetch;
    db.close();
    cleanupDbPath(dbPath);
  }
});
