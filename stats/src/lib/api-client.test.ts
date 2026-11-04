import assert from 'node:assert/strict';
import test from 'node:test';
import { apiClient, BASE_URL } from './api-client';

interface SeenRequest {
  url: string;
  method: string;
  body: unknown;
}

/** Stubs `fetch` for one test body; every call is recorded in `requests`. */
async function withFetchStub(
  respond: () => Response,
  run: (requests: SeenRequest[]) => Promise<void>,
): Promise<void> {
  const originalFetch = globalThis.fetch;
  const requests: SeenRequest[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    requests.push({
      url: String(input),
      method: init?.method ?? 'GET',
      body: init?.body ? JSON.parse(String(init.body)) : undefined,
    });
    return respond();
  }) as typeof globalThis.fetch;

  try {
    await run(requests);
  } finally {
    globalThis.fetch = originalFetch;
  }
}

const jsonResponse = (payload: unknown) =>
  new Response(JSON.stringify(payload), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });

test('getAnimeCoverUrl appends retry tokens for late cover refreshes', () => {
  const getAnimeCoverUrl = apiClient.getAnimeCoverUrl as (
    animeId: number,
    retryToken?: number,
  ) => string;

  assert.equal(getAnimeCoverUrl(42, 3), `${BASE_URL}/api/stats/anime/42/cover?coverRetry=3`);
});

test('getCoverImages dedupes anime and media ids before batching the request', async () => {
  await withFetchStub(
    () => jsonResponse({ anime: {}, media: {} }),
    async (requests) => {
      await apiClient.getCoverImages({ animeIds: [1, 1, 2], videoIds: [7, 7, 8] });
      assert.deepEqual(requests, [
        {
          url: `${BASE_URL}/api/stats/covers`,
          method: 'POST',
          body: { animeIds: [1, 2], videoIds: [7, 8] },
        },
      ]);
    },
  );
});

test('searchSentences encodes realtime sentence search requests', async () => {
  await withFetchStub(
    () => jsonResponse([]),
    async (requests) => {
      await apiClient.searchSentences('猫 食べる', 25);
      await apiClient.searchSentences('猫 食べる', 25, false);
      const query = 'q=%E7%8C%AB+%E9%A3%9F%E3%81%B9%E3%82%8B&limit=25';
      assert.deepEqual(
        requests.map((request) => request.url),
        [
          `${BASE_URL}/api/stats/sentences/search?${query}&headword=true`,
          `${BASE_URL}/api/stats/sentences/search?${query}&headword=false`,
        ],
      );
    },
  );
});

test('deleteSession throws when the stats API delete request fails', async () => {
  await withFetchStub(
    () => new Response('boom', { status: 500, statusText: 'Internal Server Error' }),
    async () => {
      await assert.rejects(() => apiClient.deleteSession(7), /Stats API error: 500 boom/);
    },
  );
});

test('getTrendsDashboard encodes a custom range and title filter', async () => {
  await withFetchStub(
    () => jsonResponse({}),
    async (requests) => {
      await apiClient.getTrendsDashboard(
        { startDay: 20_000, endDay: 20_010 },
        'week',
        true,
        'A & B',
      );
      assert.equal(
        requests[0]?.url,
        `${BASE_URL}/api/stats/trends/dashboard?range=custom&fromDay=20000&toDay=20010&groupBy=week&fillEmpty=true&title=A+%26+B`,
      );
    },
  );
});

test('getDictionaryWords sends only the filters in use', async () => {
  await withFetchStub(
    () => jsonResponse({ rows: [], total: 0, maxSeen: 0 }),
    async (requests) => {
      await apiClient.getDictionaryWords({
        search: '猫',
        sort: 'rank',
        descending: false,
        hideKnown: true,
        hideNames: false,
        hideKana: false,
        maxRank: 10_000,
        limit: 25,
        offset: 50,
      });
      const url = new URL(requests[0]!.url, 'http://stats.local');
      assert.equal(url.pathname, '/api/stats/vocabulary/words');
      assert.deepEqual(Object.fromEntries(url.searchParams), {
        sort: 'rank',
        order: 'asc',
        limit: '25',
        offset: '50',
        search: '猫',
        hideKnown: 'true',
        maxRank: '10000',
      });
    },
  );
});

test('getSessions narrows to a span of days when given one', async () => {
  await withFetchStub(
    () => jsonResponse([]),
    async (requests) => {
      await apiClient.getSessions(500, { startDay: 20_000, endDay: 20_006 });
      assert.equal(
        requests[0]?.url,
        `${BASE_URL}/api/stats/sessions?limit=500&fromDay=20000&toDay=20006`,
      );
    },
  );
});

test('getTrendsDashboard accepts 365d range and builds correct URL', async () => {
  await withFetchStub(
    () =>
      jsonResponse({
        activity: { watchTime: [], cards: [], words: [], sessions: [] },
        progress: {
          watchTime: [],
          sessions: [],
          words: [],
          newWords: [],
          cards: [],
          episodes: [],
          lookups: [],
        },
        ratios: { lookupsPerHundred: [] },
        librarySummary: [],
        animeDaily: { watchTime: [], episodes: [], cards: [], words: [] },
        patterns: { watchTimeByDayOfWeek: [], watchTimeByHour: [] },
      }),
    async (requests) => {
      await apiClient.getTrendsDashboard('365d', 'day', false);
      assert.equal(
        requests[0]?.url,
        `${BASE_URL}/api/stats/trends/dashboard?range=365d&groupBy=day&fillEmpty=false`,
      );
    },
  );
});

test('getSessionEvents can request only specific event types', async () => {
  await withFetchStub(
    () => jsonResponse([]),
    async (requests) => {
      await apiClient.getSessionEvents(42, 120, [4, 5, 6, 7, 8, 9]);
      assert.equal(
        requests[0]?.url,
        `${BASE_URL}/api/stats/sessions/42/events?limit=120&types=4%2C5%2C6%2C7%2C8%2C9`,
      );
    },
  );
});
