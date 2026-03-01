import assert from 'node:assert/strict';
import test from 'node:test';

import { buildJimakuSubtitleFilenameFromMediaPath } from './jimaku-download-path.js';

test('buildJimakuSubtitleFilenameFromMediaPath uses media basename + ja + subtitle extension', () => {
  assert.equal(
    buildJimakuSubtitleFilenameFromMediaPath('/videos/anime.mkv', 'Subs.Release.1080p.srt'),
    'anime.ja.srt',
  );
});

test('buildJimakuSubtitleFilenameFromMediaPath falls back to .srt when subtitle name has no extension', () => {
  assert.equal(
    buildJimakuSubtitleFilenameFromMediaPath('/videos/anime.mkv', 'Subs Release'),
    'anime.ja.srt',
  );
});

test('buildJimakuSubtitleFilenameFromMediaPath supports remote media URLs', () => {
  assert.equal(
    buildJimakuSubtitleFilenameFromMediaPath(
      'https://cdn.example.org/library/Anime%20Episode%2001.mkv?token=abc',
      'anything.ass',
    ),
    'Anime Episode 01.ja.ass',
  );
});
