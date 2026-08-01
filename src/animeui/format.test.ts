import test from 'node:test';
import assert from 'node:assert/strict';
import { describeInstalled, sourceOptionLabel, summarizeSearch } from './format';
import type { AnimeBrowserSearchResult } from '../types/anime-browser';

const result = (
  entryCount: number,
  failures: AnimeBrowserSearchResult['failures'] = [],
): AnimeBrowserSearchResult => ({
  entries: Array.from({ length: entryCount }, (_unused, index) => ({
    url: `/a/${index}`,
    title: `Anime ${index}`,
    thumbnailUrl: null,
    sourceId: 's',
    sourceName: 'Source',
  })),
  hasNextPage: false,
  failures,
});

test('sourceOptionLabel omits the language for an all-language source', () => {
  assert.equal(sourceOptionLabel({ id: '1', name: 'Nyaa', lang: 'ja', pkg: 'p' }), 'Nyaa (ja)');
  assert.equal(sourceOptionLabel({ id: '2', name: 'Jellyfin', lang: 'all', pkg: 'p' }), 'Jellyfin');
});

test('summarizeSearch counts results and singularizes one', () => {
  assert.equal(summarizeSearch(result(4)), '4 results');
  assert.equal(summarizeSearch(result(1)), '1 result');
  assert.equal(summarizeSearch(result(0)), '0 results');
});

test('summarizeSearch names the sources that failed alongside the ones that answered', () => {
  const summary = summarizeSearch(
    result(6, [
      { sourceId: 'a', sourceName: 'Alpha', error: 'login required' },
      { sourceId: 'b', sourceName: 'Beta', error: 'timed out' },
    ]),
  );
  assert.equal(summary, '6 results · 2 unavailable: Alpha, Beta');
});

test('describeInstalled reports sources and languages when the extension loaded', () => {
  assert.equal(
    describeInstalled({
      pkg: 'multi',
      name: 'One, Two',
      langs: ['en', 'ja'],
      sourceCount: 2,
      error: null,
    }),
    'multi · 2 sources · en, ja',
  );
});

test('describeInstalled falls back to the package alone when nothing loaded', () => {
  assert.equal(
    describeInstalled({ pkg: 'broken', name: 'broken', langs: [], sourceCount: 0, error: 'boom' }),
    'broken',
  );
});
