import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';

import { sortPlaylistBrowserDirectoryItems } from './playlist-browser-sort';

test('sortPlaylistBrowserDirectoryItems prefers parsed season and episode order', () => {
  const root = '/library/show';
  const items = sortPlaylistBrowserDirectoryItems([
    path.join(root, 'Show - S01E10.mkv'),
    path.join(root, 'Show - S01E02.mkv'),
    path.join(root, 'Show - S01E01.mkv'),
    path.join(root, 'Show - Episode 7.mkv'),
    path.join(root, 'Show - 01x03.mkv'),
  ]);

  assert.deepEqual(
    items.map((item) => item.basename),
    [
      'Show - S01E01.mkv',
      'Show - S01E02.mkv',
      'Show - 01x03.mkv',
      'Show - Episode 7.mkv',
      'Show - S01E10.mkv',
    ],
  );
  assert.deepEqual(
    items.map((item) => item.episodeLabel),
    ['S1E1', 'S1E2', 'S1E3', 'E7', 'S1E10'],
  );
});

test('sortPlaylistBrowserDirectoryItems falls back to deterministic natural ordering', () => {
  const root = '/library/show';
  const items = sortPlaylistBrowserDirectoryItems([
    path.join(root, 'Show Part 10.mkv'),
    path.join(root, 'Show Part 2.mkv'),
    path.join(root, 'Show Part 1.mkv'),
    path.join(root, 'Show Special.mkv'),
  ]);

  assert.deepEqual(
    items.map((item) => item.basename),
    ['Show Part 1.mkv', 'Show Part 2.mkv', 'Show Part 10.mkv', 'Show Special.mkv'],
  );
  assert.deepEqual(
    items.map((item) => item.episodeLabel),
    [null, null, null, null],
  );
});
