import test from 'node:test';
import assert from 'node:assert/strict';
import {
  inferAniSkipMetadataForFile,
  parseAniSkipGuessitJson,
  parseAniSkipPayload,
  resolveAniSkipMetadataForFile,
} from './aniskip-metadata';

const LOOKUP = { episodeLength: 1440 };

function makeMockResponse(payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

const MAL_MY_SHOW = { categories: [{ items: [{ id: '1234', name: 'My Show' }] }] };

function normalizeFetchInput(input: string | URL | Request): string {
  if (typeof input === 'string') return input;
  if (input instanceof URL) return input.toString();
  return input.url;
}

async function withMockFetch(
  handler: (input: string | URL | Request) => Promise<Response>,
  fn: () => Promise<void>,
): Promise<void> {
  const original = globalThis.fetch;
  (globalThis as { fetch: typeof fetch }).fetch = (async (input: string | URL | Request) => {
    return handler(input);
  }) as typeof fetch;
  try {
    await fn();
  } finally {
    (globalThis as { fetch: typeof fetch }).fetch = original;
  }
}

test('parseAniSkipGuessitJson extracts title season and episode', () => {
  const parsed = parseAniSkipGuessitJson(
    JSON.stringify({ title: 'My Show', season: 2, episode: 7 }),
    '/tmp/My.Show.S02E07.mkv',
  );
  assert.deepEqual(parsed, {
    title: 'My Show',
    season: 2,
    episode: 7,
    source: 'guessit',
    malId: null,
    intro: null,
    outro: null,
    lookupStatus: 'lookup_failed',
  });
});

test('parseAniSkipGuessitJson prefers series over episode title', () => {
  const parsed = parseAniSkipGuessitJson(
    JSON.stringify({
      title: 'What Is This, a Picnic',
      series: 'Solo Leveling',
      season: 1,
      episode: 10,
    }),
    '/tmp/S01E10-What.Is.This,.a.Picnic-Bluray-1080p.mkv',
  );
  assert.deepEqual(parsed, {
    title: 'Solo Leveling',
    season: 1,
    episode: 10,
    source: 'guessit',
    malId: null,
    intro: null,
    outro: null,
    lookupStatus: 'lookup_failed',
  });
});

test('inferAniSkipMetadataForFile falls back to filename title when guessit unavailable', () => {
  const parsed = inferAniSkipMetadataForFile('/tmp/Another_Show_-_03.mkv', {
    commandExists: () => false,
    runGuessit: () => null,
  });
  assert.equal(parsed.title.length > 0, true);
  assert.equal(parsed.source, 'fallback');
});

test('inferAniSkipMetadataForFile falls back to anime directory title when filename is episode-only', () => {
  const parsed = inferAniSkipMetadataForFile(
    '/truenas/jellyfin/anime/Solo Leveling/Season-1/S01E10-What.Is.This,.a.Picnic-Bluray-1080p.mkv',
    {
      commandExists: () => false,
      runGuessit: () => null,
    },
  );
  assert.equal(parsed.title, 'Solo Leveling');
  assert.equal(parsed.season, 1);
  assert.equal(parsed.episode, 10);
  assert.equal(parsed.source, 'fallback');
});

test('inferAniSkipMetadataForFile handles release-group filename with trailing quality tags', () => {
  const parsed = inferAniSkipMetadataForFile(
    '/media/anime/Solo Leveling/Season 1/[SubsPlease] Solo Leveling - 01 (1080p) [ABCDEF12].mkv',
    {
      commandExists: () => false,
      runGuessit: () => null,
    },
  );
  assert.equal(parsed.title, 'Solo Leveling');
  assert.equal(parsed.season, 1);
  assert.equal(parsed.episode, 1);
  assert.equal(parsed.source, 'fallback');
});

test('parseAniSkipPayload reads op and ed from a v2 response', () => {
  // Shape captured from api.aniskip.com/v2 on 2026-11-22.
  const parsed = parseAniSkipPayload({
    found: true,
    results: [
      {
        interval: { startTime: 310.571, endTime: 400.571 },
        skipType: 'op',
        skipId: 'd4f34b6d-0547-4438-ac93-b25b577eddd5',
        episodeLength: 1443.984,
      },
      {
        interval: { startTime: 1396.006, endTime: 1434 },
        skipType: 'ed',
        skipId: '5751c57a-2bcd-4dd9-be93-58e6af0f8c81',
        episodeLength: 1434.985,
      },
    ],
    message: 'Successfully found skip times',
    statusCode: 200,
  });
  assert.deepEqual(parsed, {
    intro: { start: 310.571, end: 400.571 },
    outro: { start: 1396.006, end: 1434 },
  });
});

test('parseAniSkipPayload ignores invalid intervals and not-found responses', () => {
  assert.deepEqual(
    parseAniSkipPayload({
      found: true,
      results: [
        { skipType: 'op', interval: { startTime: 50, endTime: 20 } },
        { skipType: 'op', interval: { startTime: 0, endTime: 89.5 } },
        { skipType: 'recap', interval: { startTime: 0, endTime: 30 } },
      ],
    }),
    { intro: { start: 0, end: 89.5 }, outro: null },
  );
  assert.deepEqual(parseAniSkipPayload({ found: false, results: [] }), {
    intro: null,
    outro: null,
  });
  assert.deepEqual(parseAniSkipPayload(null), { intro: null, outro: null });
});

test('resolveAniSkipMetadataForFile queries AniSkip v2 with the episode length', async () => {
  const requested: string[] = [];
  await withMockFetch(
    async (input) => {
      const url = normalizeFetchInput(input);
      requested.push(url);
      if (url.includes('myanimelist.net/search/prefix.json')) {
        return makeMockResponse({
          categories: [
            {
              items: [
                { id: '9876', name: 'Wrong Match' },
                { id: '1234', name: 'My Show' },
              ],
            },
          ],
        });
      }
      if (url.includes('api.aniskip.com/v2/skip-times/1234/7')) {
        return makeMockResponse({
          found: true,
          results: [
            { skipType: 'op', interval: { startTime: 12.5, endTime: 54.2 } },
            { skipType: 'ed', interval: { startTime: 1300, endTime: 1390 } },
          ],
        });
      }
      throw new Error(`unexpected url: ${url}`);
    },
    async () => {
      const resolved = await resolveAniSkipMetadataForFile('/media/Anime.My.Show.S01E07.mkv', {
        episodeLength: 1423.4567,
      });
      assert.equal(resolved.malId, 1234);
      assert.deepEqual(resolved.intro, { start: 12.5, end: 54.2 });
      assert.deepEqual(resolved.outro, { start: 1300, end: 1390 });
      assert.equal(resolved.lookupStatus, 'ready');
      assert.equal(resolved.title, 'Anime My Show');
    },
  );
  assert.equal(
    requested[1],
    'https://api.aniskip.com/v2/skip-times/1234/7?types=op&types=ed&episodeLength=1423.457',
  );
});

test('resolveAniSkipMetadataForFile treats an AniSkip 404 as a definite miss', async () => {
  await withMockFetch(
    async (input) => {
      const url = normalizeFetchInput(input);
      if (url.includes('myanimelist.net')) return makeMockResponse(MAL_MY_SHOW);
      return makeMockResponse(
        { found: false, results: [], message: 'No skip times found', statusCode: 404 },
        404,
      );
    },
    async () => {
      const resolved = await resolveAniSkipMetadataForFile('/media/My.Show.S01E02.mkv', LOOKUP);
      assert.equal(resolved.malId, 1234);
      assert.equal(resolved.lookupStatus, 'missing_payload');
    },
  );
});

test('resolveAniSkipMetadataForFile reports transient failures as lookup_failed', async () => {
  const failures: Array<[string, () => Promise<Response>]> = [
    ['AniSkip rate limit', async () => makeMockResponse({ statusCode: 429 }, 429)],
    ['AniSkip server error', async () => makeMockResponse({ statusCode: 503 }, 503)],
    [
      'network error',
      async () => {
        throw new TypeError('fetch failed');
      },
    ],
  ];
  for (const [label, failAniSkip] of failures) {
    await withMockFetch(
      async (input) => {
        const url = normalizeFetchInput(input);
        if (url.includes('myanimelist.net')) return makeMockResponse(MAL_MY_SHOW);
        return failAniSkip();
      },
      async () => {
        const resolved = await resolveAniSkipMetadataForFile('/media/My.Show.S01E03.mkv', LOOKUP);
        assert.equal(resolved.lookupStatus, 'lookup_failed', label);
      },
    );
  }

  await withMockFetch(
    async () => makeMockResponse({}, 429),
    async () => {
      const resolved = await resolveAniSkipMetadataForFile('/media/My.Show.S01E03.mkv', LOOKUP);
      assert.equal(resolved.lookupStatus, 'lookup_failed', 'MAL rate limit');
    },
  );
});

test('resolveAniSkipMetadataForFile emits missing_mal_id when MAL search misses', async () => {
  await withMockFetch(
    async () => makeMockResponse({ categories: [] }),
    async () => {
      const resolved = await resolveAniSkipMetadataForFile('/media/NopeShow.S01E03.mkv', LOOKUP);
      assert.equal(resolved.malId, null);
      assert.equal(resolved.lookupStatus, 'missing_mal_id');
    },
  );
});
