import test from 'node:test';
import assert from 'node:assert/strict';
import {
  describeInstalled,
  sourceOptionLabel,
  summarizeSearch,
  describeBridgeInstall,
} from './format';
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
      versionCode: 1,
      error: null,
    }),
    'multi · 2 sources · en, ja',
  );
});

test('describeInstalled falls back to the package alone when nothing loaded', () => {
  assert.equal(
    describeInstalled({
      pkg: 'broken',
      name: 'broken',
      langs: [],
      sourceCount: 0,
      versionCode: null,
      error: 'boom',
    }),
    'broken',
  );
});

test('describeBridgeInstall says who updates the bridge', () => {
  assert.match(describeBridgeInstall(null), /not started/);
  assert.match(
    describeBridgeInstall({
      origin: 'system',
      version: 'v1.0.6.2',
      dir: '/usr/share/mangatan/extension_server',
      updateAvailable: null,
    }),
    /v1\.0\.6\.2 from \/usr\/share\/mangatan\/extension_server.*package manager/,
  );
  assert.match(
    describeBridgeInstall({
      origin: 'managed',
      version: 'v1.0.5.0',
      dir: '/home/u/.config/SubMiner/anime-bridge',
      updateAvailable: 'v1.0.6.0',
    }),
    /v1\.0\.6\.0 is available/,
  );
});

test('describeBridgeInstall does not treat an unchecked managed bridge as up to date', () => {
  const description = describeBridgeInstall({
    origin: 'managed',
    version: null,
    dir: '/d',
    updateAvailable: null,
  });

  assert.match(description, /unknown version.*checks this installation for updates after startup/);
  assert.doesNotMatch(description, /up to date/);
});
