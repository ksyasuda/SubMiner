import assert from 'node:assert/strict';
import test from 'node:test';
import { apiClient, BASE_URL, resolveStatsBaseUrl } from './api-client';

test('resolveStatsBaseUrl prefers apiBase query parameter for file-based overlay mode', () => {
  const baseUrl = resolveStatsBaseUrl({
    protocol: 'file:',
    origin: 'null',
    search: '?overlay=1&apiBase=http%3A%2F%2F127.0.0.1%3A6123',
  });

  assert.equal(baseUrl, 'http://127.0.0.1:6123');
});

test('resolveStatsBaseUrl falls back to configured window origin for browser mode', () => {
  const baseUrl = resolveStatsBaseUrl({
    protocol: 'http:',
    origin: 'http://127.0.0.1:6123',
    search: '',
  });

  assert.equal(baseUrl, 'http://127.0.0.1:6123');
});

test('resolveStatsBaseUrl keeps legacy localhost fallback for file mode without apiBase', () => {
  const baseUrl = resolveStatsBaseUrl({
    protocol: 'file:',
    origin: 'null',
    search: '?overlay=1',
  });

  assert.equal(baseUrl, 'http://127.0.0.1:6969');
});

test('getAnimeCoverUrl appends retry tokens for late cover refreshes', () => {
  const getAnimeCoverUrl = apiClient.getAnimeCoverUrl as (
    animeId: number,
    retryToken?: number,
  ) => string;

  assert.equal(
    getAnimeCoverUrl(42, 3),
    'http://127.0.0.1:6969/api/stats/anime/42/cover?coverRetry=3',
  );
});

test('getCoverImages batches anime and media cover requests', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  let seenMethod = '';
  let seenBody = '';
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    seenUrl = String(input);
    seenMethod = init?.method ?? 'GET';
    seenBody = String(init?.body ?? '');
    return new Response(JSON.stringify({ anime: {}, media: {} }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as typeof globalThis.fetch;

  try {
    await apiClient.getCoverImages({ animeIds: [1, 1, 2], videoIds: [7, 7, 8] });
    assert.equal(seenUrl, `${BASE_URL}/api/stats/covers`);
    assert.equal(seenMethod, 'POST');
    assert.deepEqual(JSON.parse(seenBody), { animeIds: [1, 2], videoIds: [7, 8] });
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('deleteSession sends a DELETE request to the session endpoint', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  let seenMethod = '';
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    seenUrl = String(input);
    seenMethod = init?.method ?? 'GET';
    return new Response(null, { status: 200 });
  }) as typeof globalThis.fetch;

  try {
    await apiClient.deleteSession(42);
    assert.equal(seenUrl, `${BASE_URL}/api/stats/sessions/42`);
    assert.equal(seenMethod, 'DELETE');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('searchSentences encodes realtime sentence search requests', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    seenUrl = String(input);
    return new Response(JSON.stringify([]), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as typeof globalThis.fetch;

  try {
    await apiClient.searchSentences('猫 食べる', 25);
    assert.equal(
      seenUrl,
      `${BASE_URL}/api/stats/sentences/search?q=%E7%8C%AB+%E9%A3%9F%E3%81%B9%E3%82%8B&limit=25&headword=true`,
    );

    await apiClient.searchSentences('猫 食べる', 25, false);
    assert.equal(
      seenUrl,
      `${BASE_URL}/api/stats/sentences/search?q=%E7%8C%AB+%E9%A3%9F%E3%81%B9%E3%82%8B&limit=25&headword=false`,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('deleteSession throws when the stats API delete request fails', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response('boom', {
      status: 500,
      statusText: 'Internal Server Error',
    })) as typeof globalThis.fetch;

  try {
    await assert.rejects(() => apiClient.deleteSession(7), /Stats API error: 500 boom/);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getTrendsDashboard requests the chart-ready trends endpoint with range and grouping', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    seenUrl = String(input);
    return new Response(
      JSON.stringify({
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
        animeCumulative: {
          watchTime: [],
          episodes: [],
          cards: [],
          words: [],
        },
        patterns: {
          watchTimeByDayOfWeek: [],
          watchTimeByHour: [],
        },
      }),
      { status: 200, headers: { 'Content-Type': 'application/json' } },
    );
  }) as typeof globalThis.fetch;

  try {
    await apiClient.getTrendsDashboard('90d', 'month');
    assert.equal(
      seenUrl,
      `${BASE_URL}/api/stats/trends/dashboard?range=90d&groupBy=month&fillEmpty=true`,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getTrendsDashboard accepts 365d range and builds correct URL', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    seenUrl = String(input);
    return new Response(
      JSON.stringify({
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
        animeCumulative: {
          watchTime: [],
          episodes: [],
          cards: [],
          words: [],
        },
        patterns: {
          watchTimeByDayOfWeek: [],
          watchTimeByHour: [],
        },
      }),
      { status: 200, headers: { 'Content-Type': 'application/json' } },
    );
  }) as typeof globalThis.fetch;

  try {
    await apiClient.getTrendsDashboard('365d', 'day', false);
    assert.equal(
      seenUrl,
      `${BASE_URL}/api/stats/trends/dashboard?range=365d&groupBy=day&fillEmpty=false`,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getSessionEvents can request only specific event types', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    seenUrl = String(input);
    return new Response(JSON.stringify([]), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as typeof globalThis.fetch;

  try {
    await apiClient.getSessionEvents(42, 120, [4, 5, 6, 7, 8, 9]);
    assert.equal(
      seenUrl,
      `${BASE_URL}/api/stats/sessions/42/events?limit=120&types=4%2C5%2C6%2C7%2C8%2C9`,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getExcludedWords requests database-backed exclusions', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    seenUrl = String(input);
    return new Response(JSON.stringify([{ headword: '猫', word: '猫', reading: 'ねこ' }]), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as typeof globalThis.fetch;

  try {
    const words = await apiClient.getExcludedWords();
    assert.equal(seenUrl, `${BASE_URL}/api/stats/excluded-words`);
    assert.deepEqual(words, [{ headword: '猫', word: '猫', reading: 'ねこ' }]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('setExcludedWords replaces database-backed exclusions', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  let seenMethod = '';
  let seenBody = '';
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    seenUrl = String(input);
    seenMethod = init?.method ?? 'GET';
    seenBody = String(init?.body ?? '');
    return new Response(null, { status: 200 });
  }) as typeof globalThis.fetch;

  try {
    await apiClient.setExcludedWords([{ headword: '猫', word: '猫', reading: 'ねこ' }]);
    assert.equal(seenUrl, `${BASE_URL}/api/stats/excluded-words`);
    assert.equal(seenMethod, 'PUT');
    assert.deepEqual(JSON.parse(seenBody), {
      words: [{ headword: '猫', word: '猫', reading: 'ねこ' }],
    });
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('getSessionTimeline requests full session history when limit is omitted', async () => {
  const originalFetch = globalThis.fetch;
  let seenUrl = '';
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    seenUrl = String(input);
    return new Response(JSON.stringify([]), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as typeof globalThis.fetch;

  try {
    await apiClient.getSessionTimeline(42);
    assert.equal(seenUrl, `${BASE_URL}/api/stats/sessions/42/timeline`);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
