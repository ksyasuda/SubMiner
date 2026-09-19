import assert from 'node:assert/strict';
import test from 'node:test';
import { TmdbApiKeyMissingError, createTmdbClient, resolveTmdbApiKey } from './tmdb-client.js';

function jsonResponse(payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

function captureFetch(handler: (url: URL, init?: RequestInit) => Response) {
  const calls: Array<{ url: URL; init?: RequestInit }> = [];
  const fetchImpl = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(String(input));
    calls.push({ url, init });
    return handler(url, init);
  }) as typeof fetch;
  return { calls, fetchImpl };
}

test('resolveTmdbApiKey prefers the literal key and trims it', async () => {
  assert.equal(await resolveTmdbApiKey({ apiKey: '  abc  ', apiKeyCommand: 'echo nope' }), 'abc');
  assert.equal(await resolveTmdbApiKey({ apiKey: '', apiKeyCommand: '' }), null);
  assert.equal(await resolveTmdbApiKey(undefined), null);
});

test('resolveTmdbApiKey runs apiKeyCommand when no literal key is set', async () => {
  assert.equal(await resolveTmdbApiKey({ apiKeyCommand: 'printf " from-cmd "' }), 'from-cmd');
  assert.equal(await resolveTmdbApiKey({ apiKeyCommand: 'exit 3' }), null);
});

test('resolveTmdbApiKey falls back to the bundled key only when the user set nothing usable', async () => {
  assert.equal(await resolveTmdbApiKey({}, 'bundled'), 'bundled');
  assert.equal(await resolveTmdbApiKey({ apiKey: 'mine' }, 'bundled'), 'mine');
  assert.equal(await resolveTmdbApiKey({ apiKeyCommand: 'printf mine' }, 'bundled'), 'mine');
  assert.equal(await resolveTmdbApiKey({ apiKeyCommand: 'exit 3' }, 'bundled'), 'bundled');
});

test('search rejects without a key and never touches the network', async () => {
  const { calls, fetchImpl } = captureFetch(() => jsonResponse({ results: [] }));
  const client = createTmdbClient({ resolveApiKey: async () => null, fetch: fetchImpl });
  await assert.rejects(client.search('半沢直樹'), TmdbApiKeyMissingError);
  assert.equal(calls.length, 0);
});

test('search sends a v3 key as a query parameter and drops people from multi results', async () => {
  const { calls, fetchImpl } = captureFetch(() =>
    jsonResponse({
      results: [
        { media_type: 'person', id: 1, name: 'Sakai Masato' },
        {
          media_type: 'tv',
          id: 61222,
          name: 'Hanzawa Naoki',
          original_name: '半沢直樹',
          original_language: 'ja',
          overview: 'A banker fights back.',
          poster_path: '/hanzawa.jpg',
          first_air_date: '2013-07-07',
          genre_ids: [18],
        },
        {
          media_type: 'movie',
          id: 9,
          title: 'Anime Film',
          original_title: 'アニメ映画',
          original_language: 'ja',
          release_date: '2020-01-01',
          genre_ids: [16],
        },
      ],
    }),
  );
  const client = createTmdbClient({ resolveApiKey: async () => 'v3key', fetch: fetchImpl });

  const results = await client.search(' 半沢直樹 ');

  assert.equal(calls.length, 1);
  const url = calls[0]!.url;
  assert.equal(url.pathname, '/3/search/multi');
  assert.equal(url.searchParams.get('query'), '半沢直樹');
  assert.equal(url.searchParams.get('api_key'), 'v3key');
  assert.equal((calls[0]!.init?.headers as Record<string, string>).Authorization, undefined);
  assert.deepEqual(results, [
    {
      tmdbId: 61222,
      tmdbType: 'tv',
      title: 'Hanzawa Naoki',
      originalTitle: '半沢直樹',
      originalLanguage: 'ja',
      overview: 'A banker fights back.',
      posterUrl: 'https://image.tmdb.org/t/p/w500/hanzawa.jpg',
      year: 2013,
      isAnimation: false,
    },
    {
      tmdbId: 9,
      tmdbType: 'movie',
      title: 'Anime Film',
      originalTitle: 'アニメ映画',
      originalLanguage: 'ja',
      overview: null,
      posterUrl: null,
      year: 2020,
      isAnimation: true,
    },
  ]);
});

test('a v4 read token travels as a bearer header instead of api_key', async () => {
  const { calls, fetchImpl } = captureFetch(() => jsonResponse({ results: [] }));
  const client = createTmdbClient({
    resolveApiKey: async () => 'eyJhbGciOiJIUzI1NiJ9.payload.sig',
    fetch: fetchImpl,
  });
  await client.search('x');
  assert.equal(calls[0]!.url.searchParams.has('api_key'), false);
  assert.equal(
    (calls[0]!.init?.headers as Record<string, string>).Authorization,
    'Bearer eyJhbGciOiJIUzI1NiJ9.payload.sig',
  );
});

test('getDetails folds translations and alternative titles into the normalized shape', async () => {
  const { calls, fetchImpl } = captureFetch(() =>
    jsonResponse({
      id: 61222,
      name: 'Hanzawa Naoki',
      original_name: '半沢直樹',
      original_language: 'ja',
      overview: 'A banker fights back.',
      poster_path: '/hanzawa.jpg',
      first_air_date: '2013-07-07',
      number_of_episodes: 10,
      genres: [{ id: 18, name: 'Drama' }],
      alternative_titles: { results: [{ iso_3166_1: 'JP', title: 'Hanzawa Naoki Season 1' }] },
      translations: {
        translations: [
          { iso_639_1: 'en', data: { name: 'Hanzawa Naoki', overview: 'A banker fights back.' } },
          { iso_639_1: 'ja', data: { name: '半沢直樹', overview: '銀行員の物語' } },
        ],
      },
    }),
  );
  const client = createTmdbClient({ resolveApiKey: async () => 'k', fetch: fetchImpl });

  const details = await client.getDetails('tv', 61222);

  assert.equal(calls[0]!.url.pathname, '/3/tv/61222');
  assert.equal(
    calls[0]!.url.searchParams.get('append_to_response'),
    'alternative_titles,translations',
  );
  assert.deepEqual(details, {
    tmdbId: 61222,
    tmdbType: 'tv',
    titleEnglish: 'Hanzawa Naoki',
    titleNative: '半沢直樹',
    description: 'A banker fights back.',
    posterUrl: 'https://image.tmdb.org/t/p/w500/hanzawa.jpg',
    episodesTotal: 10,
    year: 2013,
    originalLanguage: 'ja',
    isAnimation: false,
    allTitles: ['Hanzawa Naoki', '半沢直樹', 'Hanzawa Naoki Season 1'],
  });
});

test('getDetails falls back to the Japanese overview and counts a movie as one episode', async () => {
  const { fetchImpl } = captureFetch(() =>
    jsonResponse({
      id: 5,
      title: '半沢直樹',
      original_title: '半沢直樹',
      original_language: 'ja',
      overview: '',
      release_date: '2019-03-01',
      translations: {
        translations: [{ iso_639_1: 'ja', data: { title: '半沢直樹', overview: 'あらすじ' } }],
      },
    }),
  );
  const client = createTmdbClient({ resolveApiKey: async () => 'k', fetch: fetchImpl });
  const details = await client.getDetails('movie', 5);
  assert.equal(details?.titleEnglish, null);
  assert.equal(details?.titleNative, '半沢直樹');
  assert.equal(details?.description, 'あらすじ');
  assert.equal(details?.episodesTotal, 1);
});

test('getDetails returns null for an unknown id', async () => {
  const { fetchImpl } = captureFetch(() => jsonResponse({ status_message: 'nope' }, 404));
  const client = createTmdbClient({ resolveApiKey: async () => 'k', fetch: fetchImpl });
  assert.equal(await client.getDetails('tv', 1), null);
});
