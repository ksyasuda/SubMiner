import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildLoadfileOptions,
  buildPlaybackCommands,
  buildQueuedLoadfileOptions,
  buildQueuedPlaybackCommands,
  buildTrackCommands,
  normalizeLangTag,
  selectPreferredStream,
} from './mpv-playback';
import type { ResolvedStream } from './types';

function stream(overrides: Partial<ResolvedStream> = {}): ResolvedStream {
  return {
    url: 'http://127.0.0.1:8080/video/token',
    quality: '1080p',
    headers: {},
    subtitles: [],
    audios: [],
    ...overrides,
  };
}

test('loadfile options keep one visible track and never scan the filesystem', () => {
  const options = buildLoadfileOptions({ stream: stream() });
  for (const expected of [
    'sub-auto=no',
    'secondary-sid=no',
    'secondary-sub-visibility=no',
    'sub-visibility=yes',
  ]) {
    assert.ok(options.split(',').includes(expected), `missing ${expected}`);
  }
  // The source's own subtitles are the only ones this path gets, so they must
  // not be suppressed the way the Jellyfin path suppresses them.
  assert.ok(!options.split(',').includes('sid=no'));
  // Comma-separated values would split the option list, so the language
  // preferences ride as properties instead.
  assert.ok(!options.includes('alang'));
  assert.ok(!options.includes('slang'));
});

test('queued playback carries file-local title and language preferences into mpv', () => {
  const options = {
    stream: {
      url: 'https://video.example/episode.m3u8',
      quality: '1080p',
      headers: { Referer: 'https://source.example/watch?a=1,b=2' },
      audios: [],
      subtitles: [],
    },
    title: 'Series, Part 1 = Episode 2',
  };

  const languagePreference = 'ja,jpn,jp,japanese';
  assert.ok(
    buildQueuedLoadfileOptions(options).includes(
      `alang=%${Buffer.byteLength(languagePreference)}%${languagePreference}`,
    ),
  );
  assert.ok(
    buildQueuedLoadfileOptions(options).includes(
      `force-media-title=%${Buffer.byteLength(options.title)}%${options.title}`,
    ),
  );
  assert.deepEqual(buildQueuedPlaybackCommands(options), [
    [
      'loadfile',
      'https://video.example/episode.m3u8',
      'append-play',
      -1,
      buildQueuedLoadfileOptions(options),
    ],
  ]);
});

test('headers are percent-escaped so their commas do not split the option list', () => {
  const headers = { Referer: 'https://a.test/', 'User-Agent': 'X' };
  const options = buildLoadfileOptions({ stream: stream({ headers }) });

  // Verified against mpv 0.41: the unescaped form yields an empty header list.
  const fields = 'Referer: https://a.test/,User-Agent: X';
  assert.ok(options.includes(`http-header-fields=%${fields.length}%${fields}`));
});

test('the escape length counts the full header string including separators', () => {
  const options = buildLoadfileOptions({
    stream: stream({ headers: { Cookie: 'a=1, b=2' } }),
  });
  // The comma inside the value is backslash-escaped first, so the length grows.
  const fields = 'Cookie: a=1\\, b=2';
  assert.ok(options.includes(`%${fields.length}%${fields}`));
});

test('the escape length is counted in utf-8 bytes, not js string units', () => {
  // An extension may put a non-ASCII value in a header; mpv reads %n% as a
  // byte count, so counting string units would truncate the value.
  const options = buildLoadfileOptions({
    stream: stream({ headers: { 'X-Title': '日本語' } }),
  });
  const fields = 'X-Title: 日本語';
  assert.ok(options.includes(`%${Buffer.byteLength(fields, 'utf8')}%${fields}`));
  assert.ok(!options.includes(`%${fields.length}%`));
});

test('no header option is emitted when the stream carries no headers', () => {
  const options = buildLoadfileOptions({ stream: stream() });
  assert.ok(!options.includes('http-header-fields'));
});

test('a positive start position is appended, zero is omitted', () => {
  assert.ok(buildLoadfileOptions({ stream: stream(), startSeconds: 42 }).includes('start=42'));
  assert.ok(!buildLoadfileOptions({ stream: stream(), startSeconds: 0 }).includes('start='));
  assert.ok(!buildLoadfileOptions({ stream: stream() }).includes('start='));
});

test('playback commands set the language preference before loading the file', () => {
  const commands = buildPlaybackCommands({ stream: stream(), title: 'Example - 01' });

  assert.deepEqual(commands[0], ['script-message', 'subminer-managed-subtitles-loading']);
  // Japanese first, so a multi-audio stream never starts on the dub. slang is
  // Japanese-only: an English track belongs in the secondary slot, which the
  // secondarySub auto-load fills by language tag.
  assert.deepEqual(commands[1], ['set_property', 'alang', 'ja,jpn,jp,japanese']);
  assert.deepEqual(commands[2], ['set_property', 'slang', 'ja,jpn,jp,japanese']);
  assert.equal(commands[3]?.[0], 'loadfile');
  assert.equal(commands[3]?.[1], 'http://127.0.0.1:8080/video/token');
  assert.equal(commands[3]?.[2], 'replace');
  assert.equal(commands[3]?.[3], -1);
  assert.deepEqual(commands[4], ['set_property', 'force-media-title', 'Example - 01']);
});

test('force-media-title is skipped when there is no title', () => {
  assert.equal(buildPlaybackCommands({ stream: stream() }).length, 4);
  assert.equal(buildPlaybackCommands({ stream: stream(), title: '' }).length, 4);
});

test('external audio tracks are added, with the Japanese one selected', () => {
  const commands = buildTrackCommands(
    stream({
      audios: [
        { url: 'http://host/en.m4a', lang: 'en' },
        { url: 'http://host/ja.m4a', lang: 'ja' },
      ],
    }),
  );

  assert.deepEqual(commands, [
    ['audio-add', 'http://host/en.m4a', 'auto', 'en', 'en'],
    ['audio-add', 'http://host/ja.m4a', 'select', 'ja', 'ja'],
  ]);
});

test('external audio is left unselected when none of it is Japanese', () => {
  // alang already picked a track off the container; do not override it.
  const commands = buildTrackCommands(
    stream({ audios: [{ url: 'http://host/en.m4a', lang: 'eng' }] }),
  );

  assert.deepEqual(commands, [['audio-add', 'http://host/en.m4a', 'auto', 'eng', 'en']]);
});

test('only a Japanese subtitle track is selected as primary', () => {
  const japanese = buildTrackCommands(
    stream({
      subtitles: [
        { url: 'http://host/en.vtt', lang: 'English' },
        { url: 'http://host/ja.vtt', lang: 'Japanese' },
      ],
    }),
  );
  assert.deepEqual(japanese[1], ['sub-add', 'http://host/ja.vtt', 'select', 'Japanese', 'ja']);
  assert.equal(japanese[0]?.[2], 'auto');

  // English is the user's *secondary* language; it must not take the primary
  // slot. It rides in unselected, tagged so the secondarySub auto-load can
  // route it to secondary-sid.
  const englishOnly = buildTrackCommands(
    stream({ subtitles: [{ url: 'http://host/en.vtt', lang: 'English' }] }),
  );
  assert.deepEqual(englishOnly, [['sub-add', 'http://host/en.vtt', 'auto', 'English', 'en']]);
});

test('language labels normalize to the tags users configure', () => {
  assert.equal(normalizeLangTag('English'), 'en');
  assert.equal(normalizeLangTag('eng'), 'en');
  assert.equal(normalizeLangTag('en-US'), 'en');
  assert.equal(normalizeLangTag('Japanese'), 'ja');
  assert.equal(normalizeLangTag('jpn'), 'ja');
  assert.equal(normalizeLangTag('Português'), 'pt');
  // Unknown labels pass through untouched rather than being guessed at.
  assert.equal(normalizeLangTag('Klingon'), 'Klingon');
  assert.equal(normalizeLangTag(''), '');
});

test('unlabelled tracks still get a usable menu title, duplicates are dropped', () => {
  const commands = buildTrackCommands(
    stream({
      subtitles: [
        { url: 'http://host/a.vtt', lang: '' },
        { url: 'http://host/a.vtt', lang: '' },
        { url: 'http://host/b.vtt', lang: '' },
      ],
    }),
  );

  assert.deepEqual(commands, [
    ['sub-add', 'http://host/a.vtt', 'auto', 'Subtitle 1', ''],
    ['sub-add', 'http://host/b.vtt', 'auto', 'Subtitle 2', ''],
  ]);
});

test('a stream with no external tracks emits no track commands', () => {
  assert.deepEqual(buildTrackCommands(stream()), []);
});

test('selectPreferredStream skips dub entries in favour of the original audio', () => {
  const streams = [
    stream({ quality: '1080p (Dub)' }),
    stream({ quality: '720p (Sub)' }),
    stream({ quality: '480p (Dub)' }),
  ];

  // Language beats the quality hint: a 1080p dub is the wrong file, not a
  // better one.
  assert.equal(selectPreferredStream(streams)?.quality, '720p (Sub)');
  assert.equal(selectPreferredStream(streams, '1080')?.quality, '720p (Sub)');
});

test('selectPreferredStream prefers an entry carrying a Japanese audio track', () => {
  const streams = [
    stream({ quality: '1080p' }),
    stream({ quality: '720p', audios: [{ url: 'http://host/ja.m4a', lang: 'ja' }] }),
  ];

  assert.equal(selectPreferredStream(streams)?.quality, '720p');
});

test('an all-dub list still plays rather than failing', () => {
  const streams = [stream({ quality: '1080p Dub' }), stream({ quality: '720p Dub' })];

  assert.equal(selectPreferredStream(streams)?.quality, '1080p Dub');
  assert.equal(selectPreferredStream(streams, '720')?.quality, '720p Dub');
});

test('selectPreferredStream honours a quality hint, else takes the first', () => {
  const streams = [stream({ quality: '360p' }), stream({ quality: '1080p' })];

  assert.equal(selectPreferredStream(streams, '1080')?.quality, '1080p');
  assert.equal(selectPreferredStream(streams, '1080P')?.quality, '1080p');
  // Extensions label streams with the host name, so the hint matches a substring.
  const decorated = [
    stream({ quality: 'Doodstream - 360p' }),
    stream({ quality: 'Vidhide - 720p' }),
  ];
  assert.equal(selectPreferredStream(decorated, '720')?.quality, 'Vidhide - 720p');
  // Extensions pre-sort by their own preference, so the first entry wins.
  assert.equal(selectPreferredStream(streams)?.quality, '360p');
  // A hint that matches nothing falls back rather than failing.
  assert.equal(selectPreferredStream(streams, '4k')?.quality, '360p');
});

test('selectPreferredStream returns null for an empty list', () => {
  assert.equal(selectPreferredStream([]), null);
  assert.equal(selectPreferredStream([], '1080p'), null);
});
