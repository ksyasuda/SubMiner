import assert from 'node:assert/strict';
import test from 'node:test';

import { parseMediaInfo, parseMediaLookupInfo, parseMediaTitleInfo } from './utils.js';

const JELLYFIN_TITLE =
  "Seitokai Yakuindomo S01E01 Under the Cherry Tree / Will I Get This Feeling Every Time? / For Now, Let's Try Stripping";

test('parseMediaTitleInfo keeps the series when the episode name contains slashes', () => {
  const info = parseMediaTitleInfo(JELLYFIN_TITLE);
  assert.equal(info.title, 'Seitokai Yakuindomo');
  assert.equal(info.season, 1);
  assert.equal(info.episode, 1);
  assert.equal(info.confidence, 'high');
});

test('parseMediaTitleInfo does not treat a dotted title as a file extension', () => {
  // Dots are normalized to spaces; the point is that "Stone S02E03" is not cut off as an extension.
  assert.equal(parseMediaTitleInfo('Dr. Stone S02E03').title, 'Dr Stone');
  assert.equal(parseMediaTitleInfo('Dr. Stone S02E03.mkv').episode, 3);
});

test('parseMediaInfo still reads the season from the parent directory', () => {
  const info = parseMediaInfo('/media/Show/Season 2/Show - 03 - Title.mkv');
  assert.equal(info.title, 'Show');
  assert.equal(info.season, 2);
  assert.equal(info.episode, 3);
});

test('parseMediaLookupInfo parses a stream by its title, not its URL', () => {
  const info = parseMediaLookupInfo('jellyfin://host:8096/item/abc', JELLYFIN_TITLE);
  assert.equal(info.title, 'Seitokai Yakuindomo');
  assert.equal(info.episode, 1);
});

test('parseMediaLookupInfo treats a title already swapped in for the path as a title', () => {
  assert.equal(parseMediaLookupInfo(JELLYFIN_TITLE, JELLYFIN_TITLE).title, 'Seitokai Yakuindomo');
});
