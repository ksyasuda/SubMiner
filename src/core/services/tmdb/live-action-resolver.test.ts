import assert from 'node:assert/strict';
import test from 'node:test';
import { createLiveActionMetadataResolver, titlesMatch } from './live-action-resolver.js';
import {
  TmdbApiKeyMissingError,
  type TmdbClient,
  type TmdbSearchResult,
  type TmdbTitleDetails,
} from './tmdb-client.js';

const silentLogger = { info: () => {}, warn: () => {} };

function searchResult(over: Partial<TmdbSearchResult> & { tmdbId: number }): TmdbSearchResult {
  return {
    tmdbType: 'tv',
    title: 'Title',
    originalTitle: 'Title',
    originalLanguage: 'ja',
    overview: null,
    posterUrl: null,
    year: null,
    isAnimation: false,
    ...over,
  };
}

function details(over: Partial<TmdbTitleDetails> & { tmdbId: number }): TmdbTitleDetails {
  return {
    tmdbType: 'tv',
    titleEnglish: null,
    titleNative: null,
    description: null,
    posterUrl: null,
    episodesTotal: null,
    year: null,
    originalLanguage: 'ja',
    isAnimation: false,
    allTitles: [],
    ...over,
  };
}

test('titlesMatch ignores case, width, and punctuation but not extra words', () => {
  assert.equal(titlesMatch('Hanzawa Naoki', ['HANZAWA　NAOKI!']), true);
  assert.equal(titlesMatch('半沢直樹', ['半沢直樹']), true);
  assert.equal(titlesMatch('Hanzawa Naoki', ['Hanzawa Naoki Season 2']), false);
  assert.equal(titlesMatch('', ['']), false);
});

test('resolveByTitle only accepts a Japanese non-animated result whose known titles match exactly', async () => {
  const detailCalls: number[] = [];
  const client: TmdbClient = {
    async search() {
      return [
        searchResult({ tmdbId: 1, title: 'Hanzawa Naoki', originalLanguage: 'ko' }),
        searchResult({ tmdbId: 2, title: 'Hanzawa Naoki', isAnimation: true }),
        searchResult({ tmdbId: 3, title: 'Hanzawa Naoki: The Movie' }),
        searchResult({ tmdbId: 4, title: 'Hanzawa Naoki' }),
      ];
    },
    async getDetails(_type, tmdbId) {
      detailCalls.push(tmdbId);
      if (tmdbId === 3) return details({ tmdbId: 3, allTitles: ['Hanzawa Naoki: The Movie'] });
      if (tmdbId === 4) return details({ tmdbId: 4, allTitles: ['Hanzawa Naoki', '半沢直樹'] });
      return null;
    },
  };
  const resolver = createLiveActionMetadataResolver(client, silentLogger);

  const resolved = await resolver.resolveByTitle('hanzawa naoki');

  assert.equal(resolved?.tmdbId, 4);
  assert.deepEqual(detailCalls, [3, 4]);
});

test('resolveByTitle returns null when nothing matches or the key is missing', async () => {
  const noMatch: TmdbClient = {
    async search() {
      return [searchResult({ tmdbId: 1, title: 'Something Else' })];
    },
    async getDetails() {
      return details({ tmdbId: 1, allTitles: ['Something Else'] });
    },
  };
  assert.equal(
    await createLiveActionMetadataResolver(noMatch, silentLogger).resolveByTitle('Hanzawa Naoki'),
    null,
  );

  let infoCount = 0;
  const noKey: TmdbClient = {
    async search() {
      throw new TmdbApiKeyMissingError();
    },
    async getDetails() {
      throw new TmdbApiKeyMissingError();
    },
  };
  const resolver = createLiveActionMetadataResolver(noKey, {
    info: () => {
      infoCount += 1;
    },
    warn: () => {},
  });
  assert.equal(await resolver.resolveByTitle('Hanzawa Naoki'), null);
  assert.equal(await resolver.resolveById('tv', 1), null);
  assert.equal(infoCount, 1);
});
