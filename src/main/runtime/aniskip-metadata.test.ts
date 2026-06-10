import test from 'node:test';
import assert from 'node:assert/strict';
import {
  inferAniSkipMetadataForFile,
  parseAniSkipGuessitJson,
  resolveAniSkipMetadataForFile,
} from './aniskip-metadata';

function makeMockResponse(payload: unknown): Response {
  return {
    ok: true,
    status: 200,
    json: async () => payload,
  } as Response;
}

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
    introStart: null,
    introEnd: null,
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
    introStart: null,
    introEnd: null,
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

test('resolveAniSkipMetadataForFile resolves MAL id and intro payload', async () => {
  await withMockFetch(
    async (input) => {
      const url = normalizeFetchInput(input);
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
      if (url.includes('api.aniskip.com/v1/skip-times/1234/7')) {
        return makeMockResponse({
          found: true,
          results: [{ skip_type: 'op', interval: { start_time: 12.5, end_time: 54.2 } }],
        });
      }
      throw new Error(`unexpected url: ${url}`);
    },
    async () => {
      const resolved = await resolveAniSkipMetadataForFile('/media/Anime.My.Show.S01E07.mkv');
      assert.equal(resolved.malId, 1234);
      assert.equal(resolved.introStart, 12.5);
      assert.equal(resolved.introEnd, 54.2);
      assert.equal(resolved.lookupStatus, 'ready');
      assert.equal(resolved.title, 'Anime My Show');
    },
  );
});

test('resolveAniSkipMetadataForFile accepts intro payload starting at zero', async () => {
  await withMockFetch(
    async (input) => {
      const url = normalizeFetchInput(input);
      if (url.includes('myanimelist.net/search/prefix.json')) {
        return makeMockResponse({
          categories: [{ items: [{ id: '1234', name: 'My Show' }] }],
        });
      }
      if (url.includes('api.aniskip.com/v1/skip-times/1234/1')) {
        return makeMockResponse({
          found: true,
          results: [{ skip_type: 'op', interval: { start_time: 0, end_time: 89.5 } }],
        });
      }
      throw new Error(`unexpected url: ${url}`);
    },
    async () => {
      const resolved = await resolveAniSkipMetadataForFile('/media/My.Show.S01E01.mkv');
      assert.equal(resolved.malId, 1234);
      assert.equal(resolved.introStart, 0);
      assert.equal(resolved.introEnd, 89.5);
      assert.equal(resolved.lookupStatus, 'ready');
    },
  );
});

test('resolveAniSkipMetadataForFile emits missing_mal_id when MAL search misses', async () => {
  await withMockFetch(
    async () => makeMockResponse({ categories: [] }),
    async () => {
      const resolved = await resolveAniSkipMetadataForFile('/media/NopeShow.S01E03.mkv');
      assert.equal(resolved.malId, null);
      assert.equal(resolved.lookupStatus, 'missing_mal_id');
    },
  );
});
