import assert from 'node:assert/strict';
import test from 'node:test';
import { animeSubtitleGenerationSources } from './anime-subtitle-generation-sources';
import type { ResolvedStream } from '../../anime-bridge/types';

test('extraction prioritizes Japanese audio then smaller video without changing playback', () => {
  const selected: ResolvedStream = {
    url: 'https://stream/high',
    quality: '1080p',
    headers: { Referer: 'https://source/', Cookie: 'private', Authorization: 'private' },
    subtitles: [],
    audios: [
      { url: 'https://stream/en', lang: 'English' },
      { url: 'https://stream/ja', lang: 'Japanese' },
    ],
  };
  const candidates = animeSubtitleGenerationSources(selected, [
    selected,
    { ...selected, url: 'https://stream/720', quality: '720p' },
    { ...selected, url: 'https://stream/dub', quality: '360p Dub' },
    { ...selected, url: 'https://stream/360', quality: '360p' },
    { ...selected, url: 'https://stream/unknown', quality: 'Auto' },
  ]);
  assert.deepEqual(
    candidates.map(({ kind, url }) => [kind, url]),
    [
      ['audio', 'https://stream/ja'],
      ['video', 'https://stream/360'],
      ['video', 'https://stream/720'],
    ],
  );
  assert.equal(selected.url, 'https://stream/high');
  assert.deepEqual(candidates[0]?.httpHeaders.headers, { Referer: 'https://source/' });
});
