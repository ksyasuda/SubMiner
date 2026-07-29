import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { createCoverArtFetcher, stripFilenameTags } from './cover-art-fetcher.js';
import { Database } from '../immersion-tracker/sqlite.js';
import {
  ensureSchema,
  getOrCreateAnimeRecord,
  getOrCreateVideoRecord,
  linkVideoToAnimeRecord,
} from '../immersion-tracker/storage.js';
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

test('fetchIfMissing reuses cached cover art from another video in the same anime', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  const firstVideoId = getOrCreateVideoRecord(db, 'local:/tmp/cover-fetcher-cache-1.mkv', {
    canonicalTitle: 'Shared Cover Show',
    sourcePath: '/tmp/cover-fetcher-cache-1.mkv',
    sourceUrl: null,
    sourceType: SOURCE_TYPE_LOCAL,
  });
  const secondVideoId = getOrCreateVideoRecord(db, 'local:/tmp/cover-fetcher-cache-2.mkv', {
    canonicalTitle: 'Shared Cover Show',
    sourcePath: '/tmp/cover-fetcher-cache-2.mkv',
    sourceUrl: null,
    sourceType: SOURCE_TYPE_LOCAL,
  });
  const animeId = getOrCreateAnimeRecord(db, {
    parsedTitle: 'Shared Cover Show',
    canonicalTitle: 'Shared Cover Show',
    anilistId: 99,
    titleRomaji: 'Shared Cover Show',
    titleEnglish: 'Shared Cover Show',
    titleNative: null,
    metadataJson: null,
  });
  for (const videoId of [firstVideoId, secondVideoId]) {
    linkVideoToAnimeRecord(db, videoId, {
      animeId,
      parsedBasename: null,
      parsedTitle: 'Shared Cover Show',
      parsedSeason: 1,
      parsedEpisode: videoId,
      parserSource: 'fallback',
      parserConfidence: 1,
      parseMetadataJson: null,
    });
  }
  upsertCoverArt(db, firstVideoId, {
    anilistId: 99,
    coverUrl: 'https://images.test/shared-cover.jpg',
    coverBlob: Buffer.from([9, 8, 7, 6]),
    titleRomaji: 'Shared Cover Show',
    titleEnglish: 'Shared Cover Show',
    episodesTotal: 12,
  });

  let fetchCalls = 0;
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () => {
    fetchCalls += 1;
    throw new Error('unexpected AniList or image request');
  }) as typeof fetch;

  try {
    const fetcher = createCoverArtFetcher(
      {
        acquire: async () => {},
        recordResponse: () => {},
      },
      console,
    );

    const fetched = await fetcher.fetchIfMissing(db, secondVideoId, 'Shared Cover Show');
    const stored = getCoverArt(db, secondVideoId);

    assert.equal(fetched, true);
    assert.equal(fetchCalls, 0);
    assert.equal(stored?.anilistId, 99);
    assert.equal(Buffer.from(stored?.coverBlob ?? []).toString('hex'), '09080706');
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
  const relationCalls: number[] = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
    const raw = (init?.body as string | undefined) ?? '';
    const payload = JSON.parse(raw) as { variables: { search?: string; id?: number } };

    if (typeof payload.variables.id === 'number') {
      relationCalls.push(payload.variables.id);
      return Promise.resolve(
        createJsonResponse({
          data: {
            Media: {
              relations: {
                edges: [
                  {
                    relationType: 'SEQUEL',
                    node: {
                      id: 20,
                      type: 'ANIME',
                      episodes: 25,
                      format: 'TV',
                      seasonYear: 2017,
                      coverImage: { large: 'https://images.test/cover-s2.jpg', medium: null },
                      title: {
                        romaji: 'Little Witch Academia 2',
                        english: 'Little Witch Academia 2',
                        native: null,
                      },
                    },
                  },
                ],
              },
            },
          },
        }),
      );
    }

    searchCalls.push({ search: String(payload.variables.search) });
    return Promise.resolve(
      createJsonResponse({
        data: {
          Page: {
            media: [
              {
                id: 19,
                episodes: 24,
                format: 'TV',
                seasonYear: 2013,
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
    // One search on the bare title, then a sequel hop to reach season 2.
    assert.equal(searchCalls.length, 1);
    assert.equal(searchCalls[0]!.search, 'Little Witch Academia');
    assert.deepEqual(relationCalls, [19]);
    assert.equal(stored?.anilistId, 20);
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

test('fetchIfMissing caches a no-match when the season cannot be resolved', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  const videoId = getOrCreateVideoRecord(db, 'local:/tmp/cover-fetcher-unresolved.mkv', {
    canonicalTitle: 'Unresolved Show (2013) - S03E01 - Something [1080p].mkv',
    sourcePath: '/tmp/cover-fetcher-unresolved.mkv',
    sourceType: SOURCE_TYPE_LOCAL,
    sourceUrl: null,
  });

  const originalFetch = globalThis.fetch;
  globalThis.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (!url.includes('graphql')) {
      return Promise.resolve(
        new Response(Buffer.from('01020304'), {
          status: 200,
          headers: { 'content-type': 'image/png' },
        }),
      );
    }

    const payload = JSON.parse(String(init?.body ?? '{}')) as {
      variables?: { search?: string; id?: number };
    };
    if (typeof payload.variables?.id === 'number') {
      // No sequel edges, so season 3 cannot be reached from the season 1 anchor.
      return Promise.resolve(createJsonResponse({ data: { Media: { relations: { edges: [] } } } }));
    }
    return Promise.resolve(
      createJsonResponse({
        data: {
          Page: {
            media: [
              {
                id: 55,
                episodes: 13,
                format: 'TV',
                seasonYear: 2013,
                coverImage: { large: 'https://images.test/s1.jpg', medium: null },
                title: { romaji: 'Unresolved Show', english: 'Unresolved Show', native: null },
              },
            ],
          },
        },
      }),
    );
  }) as typeof fetch;

  try {
    const fetcher = createCoverArtFetcher(
      { acquire: async () => {}, recordResponse: () => {} },
      console,
      {
        runGuessit: async () =>
          JSON.stringify({ title: 'Unresolved Show', season: 3, episode: 1, year: 2013 }),
      },
    );

    const fetched = await fetcher.fetchIfMissing(db, videoId, 'Unresolved Show');
    const stored = getCoverArt(db, videoId);

    assert.equal(fetched, false);
    // Storing the season 1 artwork would leave a blob with no AniList id, which the
    // `existing.coverBlob` early return serves forever - the season could then never
    // re-resolve. A plain no-match keeps the existing retry window in play instead.
    assert.equal(stored?.coverBlob, null);
    assert.equal(stored?.coverUrl, null);
    assert.equal(stored?.anilistId, null);
    assert.equal(stored?.episodesTotal, null);
  } finally {
    globalThis.fetch = originalFetch;
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('fetchIfMissing re-resolves an unresolved season once AniList publishes the relation', async () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  const videoId = getOrCreateVideoRecord(db, 'local:/tmp/cover-fetcher-recovers.mkv', {
    canonicalTitle: 'Recovering Show (2013) - S02E01 - Something [1080p].mkv',
    sourcePath: '/tmp/cover-fetcher-recovers.mkv',
    sourceType: SOURCE_TYPE_LOCAL,
    sourceUrl: null,
  });

  let sequelPublished = false;
  const originalFetch = globalThis.fetch;
  const originalNow = Date.now;
  globalThis.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (!url.includes('graphql')) {
      return Promise.resolve(
        new Response(Buffer.from('01020304'), {
          status: 200,
          headers: { 'content-type': 'image/png' },
        }),
      );
    }

    const payload = JSON.parse(String(init?.body ?? '{}')) as {
      variables?: { search?: string; id?: number };
    };
    if (typeof payload.variables?.id === 'number') {
      return Promise.resolve(
        createJsonResponse({
          data: {
            Media: {
              relations: {
                edges: sequelPublished
                  ? [
                      {
                        relationType: 'SEQUEL',
                        node: {
                          id: 66,
                          type: 'ANIME',
                          episodes: 12,
                          format: 'TV',
                          seasonYear: 2015,
                          coverImage: { large: 'https://images.test/s2.jpg', medium: null },
                          title: { romaji: 'Recovering Show 2', english: null, native: null },
                        },
                      },
                    ]
                  : [],
              },
            },
          },
        }),
      );
    }
    return Promise.resolve(
      createJsonResponse({
        data: {
          Page: {
            media: [
              {
                id: 65,
                episodes: 13,
                format: 'TV',
                seasonYear: 2013,
                coverImage: { large: 'https://images.test/s1.jpg', medium: null },
                title: { romaji: 'Recovering Show', english: null, native: null },
              },
            ],
          },
        },
      }),
    );
  }) as typeof fetch;

  try {
    const fetcher = createCoverArtFetcher(
      { acquire: async () => {}, recordResponse: () => {} },
      console,
      {
        runGuessit: async () =>
          JSON.stringify({ title: 'Recovering Show', season: 2, episode: 1, year: 2013 }),
      },
    );

    assert.equal(await fetcher.fetchIfMissing(db, videoId, 'Recovering Show'), false);
    assert.equal(getCoverArt(db, videoId)?.anilistId, null);

    // AniList publishes the sequel relation, and the no-match retry window elapses.
    sequelPublished = true;
    const base = originalNow();
    Date.now = () => base + 10 * 60 * 1000;

    assert.equal(await fetcher.fetchIfMissing(db, videoId, 'Recovering Show'), true);
    const recovered = getCoverArt(db, videoId);
    assert.equal(recovered?.anilistId, 66);
    assert.equal(recovered?.episodesTotal, 12);
  } finally {
    Date.now = originalNow;
    globalThis.fetch = originalFetch;
    db.close();
    cleanupDbPath(dbPath);
  }
});
