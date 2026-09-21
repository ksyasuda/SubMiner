import assert from 'node:assert/strict';
import test, { type TestContext } from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { TmdbConfig } from '../../../types/integrations';
import {
  TmdbApiKeyMissingError,
  createTmdbClient,
  createTmdbApiKeyResolver,
  resolveTmdbApiKey,
} from './tmdb-client.js';

function commandFixture(t: TestContext) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer tmdb command-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  let nextId = 0;
  const quotePath = (value: string) =>
    `"${process.platform === 'win32' ? value : value.replace(/["\\$`]/g, '\\$&')}"`;
  return {
    dir,
    command(source: string): string {
      const script = path.join(dir, `credential-${nextId++}.cjs`);
      fs.writeFileSync(script, source);
      return `${quotePath(process.execPath)} ${quotePath(script)}`;
    },
  };
}

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

test('resolveTmdbApiKey prefers the literal key and trims it', async (t) => {
  const { command } = commandFixture(t);
  assert.equal(
    await resolveTmdbApiKey({ apiKey: '  abc  ', apiKeyCommand: command('process.exit(3)') }),
    'abc',
  );
  assert.equal(await resolveTmdbApiKey({ apiKey: '', apiKeyCommand: '' }), null);
  assert.equal(await resolveTmdbApiKey(undefined), null);
});

test('resolveTmdbApiKey runs apiKeyCommand when no literal key is set', async (t) => {
  const { command } = commandFixture(t);
  assert.equal(
    await resolveTmdbApiKey({ apiKeyCommand: command('process.stdout.write(" from-cmd ")') }),
    'from-cmd',
  );
  assert.equal(await resolveTmdbApiKey({ apiKeyCommand: command('process.exit(3)') }), null);
});

test('resolveTmdbApiKey falls back to the bundled key only when the user set nothing usable', async (t) => {
  const { command } = commandFixture(t);
  assert.equal(await resolveTmdbApiKey({}, 'bundled'), 'bundled');
  assert.equal(await resolveTmdbApiKey({ apiKey: 'mine' }, 'bundled'), 'mine');
  assert.equal(
    await resolveTmdbApiKey({ apiKeyCommand: command('process.stdout.write("mine")') }, 'bundled'),
    'mine',
  );
  assert.equal(
    await resolveTmdbApiKey({ apiKeyCommand: command('process.exit(3)') }, 'bundled'),
    'bundled',
  );
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
  const v4Token = ['eyJ', 'test-header', '.payload', '.sig'].join('');
  const { calls, fetchImpl } = captureFetch(() => jsonResponse({ results: [] }));
  const client = createTmdbClient({
    resolveApiKey: async () => v4Token,
    fetch: fetchImpl,
  });
  await client.search('x');
  assert.equal(calls[0]!.url.searchParams.has('api_key'), false);
  assert.equal(
    (calls[0]!.init?.headers as Record<string, string>).Authorization,
    `Bearer ${v4Token}`,
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

test('client reuses command output across requests and invalidates it when either setting changes', async (t) => {
  const fixture = commandFixture(t);
  const counter = path.join(fixture.dir, 'calls');
  const createCommand = (key: string) =>
    fixture.command(`
    const fs = require('node:fs');
    const path = require('node:path');
    fs.appendFileSync(path.join(__dirname, 'calls'), 'x');
    process.stdout.write(${JSON.stringify(key)});
  `);
  let config: TmdbConfig = { apiKeyCommand: createCommand('command-key') };
  const { calls, fetchImpl } = captureFetch(() => jsonResponse({ results: [] }));
  const client = createTmdbClient({
    resolveApiKey: createTmdbApiKeyResolver(
      () => config,
      () => 'bundled',
    ),
    fetch: fetchImpl,
  });
  await Promise.all([client.search('a'), client.search('b')]);
  await client.getDetails('tv', 1);
  assert.equal(fs.readFileSync(counter, 'utf8'), 'x');
  assert.ok(calls.every(({ url }) => url.searchParams.get('api_key') === 'command-key'));
  config = { ...config, apiKey: 'literal' };
  await client.search('c');
  assert.equal(calls.at(-1)?.url.searchParams.get('api_key'), 'literal');
  config = { ...config, apiKey: '' };
  await client.search('d');
  assert.equal(fs.readFileSync(counter, 'utf8'), 'xx');
  config = { apiKeyCommand: createCommand('new-key') };
  await client.search('e');
  assert.equal(fs.readFileSync(counter, 'utf8'), 'xxx');
  assert.equal(calls.at(-1)?.url.searchParams.get('api_key'), 'new-key');
});

for (const failure of ['error', 'empty'] as const) {
  test(`${failure} command output uses a bounded cooldown before retrying`, async (t) => {
    const fixture = commandFixture(t);
    const counter = path.join(fixture.dir, 'calls');
    const command = fixture.command(`
      const fs = require('node:fs');
      const path = require('node:path');
      const counter = path.join(__dirname, 'calls');
      fs.appendFileSync(counter, 'x');
      if (fs.readFileSync(counter, 'utf8').length === 1) process.exit(${failure === 'error' ? 1 : 0});
      process.stdout.write('recovered');
    `);
    let now = 1000;
    const originalNow = Date.now;
    Date.now = () => now;
    t.after(() => {
      Date.now = originalNow;
    });
    let bundledKey: string | null = 'bundled';
    const resolve = createTmdbApiKeyResolver(
      () => ({ apiKeyCommand: command }),
      () => bundledKey,
    );
    assert.deepEqual(await Promise.all([resolve(), resolve()]), ['bundled', 'bundled']);
    now += 29_999;
    assert.equal(await resolve(), 'bundled');
    bundledKey = null;
    assert.equal(await resolve(), null);
    assert.equal(fs.readFileSync(counter, 'utf8'), 'x');
    now += 1;
    assert.equal(await resolve(), 'recovered');
    assert.equal(await resolve(), 'recovered');
    assert.equal(fs.readFileSync(counter, 'utf8'), 'xx');
  });
}

for (const setting of ['apiKey', 'apiKeyCommand'] as const) {
  test(`changing ${setting} clears a failed command cooldown`, async (t) => {
    const fixture = commandFixture(t);
    const counter = path.join(fixture.dir, 'calls');
    const source = `
      const fs = require('node:fs');
      const path = require('node:path');
      fs.appendFileSync(path.join(__dirname, 'calls'), 'x');
      process.exit(1);
    `;
    let config: TmdbConfig = { apiKeyCommand: fixture.command(source) };
    const resolve = createTmdbApiKeyResolver(
      () => config,
      () => 'bundled',
    );
    assert.equal(await resolve(), 'bundled');
    assert.equal(await resolve(), 'bundled');
    assert.equal(fs.readFileSync(counter, 'utf8'), 'x');
    config =
      setting === 'apiKey'
        ? { ...config, apiKey: ' ' }
        : { apiKeyCommand: fixture.command(source) };
    assert.equal(await resolve(), 'bundled');
    assert.equal(fs.readFileSync(counter, 'utf8'), 'xx');
  });
}
