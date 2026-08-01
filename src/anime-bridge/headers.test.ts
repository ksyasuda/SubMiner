import test from 'node:test';
import assert from 'node:assert/strict';
import { parseOkHttpHeaders, resolveStream, toMpvHeaderFields } from './headers';

test('parseOkHttpHeaders flattens the alternating name/value array', () => {
  const parsed = parseOkHttpHeaders({
    namesAndValues$okhttp: ['Referer', 'https://origin.example/', 'User-Agent', 'Aniyomi'],
  });
  assert.deepEqual(parsed, {
    Referer: 'https://origin.example/',
    'User-Agent': 'Aniyomi',
  });
});

test('parseOkHttpHeaders tolerates missing, empty, and odd-length input', () => {
  assert.deepEqual(parseOkHttpHeaders(undefined), {});
  assert.deepEqual(parseOkHttpHeaders({}), {});
  assert.deepEqual(parseOkHttpHeaders({ namesAndValues$okhttp: [] }), {});
  // A trailing name with no value is dropped rather than mapped to undefined.
  assert.deepEqual(parseOkHttpHeaders({ namesAndValues$okhttp: ['Referer'] }), {});
});

test('toMpvHeaderFields joins entries and escapes commas in values', () => {
  const fields = toMpvHeaderFields({
    Referer: 'https://origin.example/',
    Cookie: 'a=1, b=2',
  });
  assert.equal(fields, 'Referer: https://origin.example/,Cookie: a=1\\, b=2');
});

test('toMpvHeaderFields returns an empty string when there are no headers', () => {
  assert.equal(toMpvHeaderFields({}), '');
});

test('resolveStream normalizes a bridge video into a playable stream', () => {
  const stream = resolveStream({
    url: 'https://origin.example/embed/1',
    quality: '1080p',
    videoUrl: 'http://127.0.0.1:8080/video/master-token',
    headers: { namesAndValues$okhttp: ['Referer', 'https://origin.example/'] },
    subtitleTracks: [{ url: 'http://127.0.0.1:8080/video/sub-token', lang: 'English' }],
    audioTracks: [{ url: 'http://127.0.0.1:8080/video/audio-token', lang: 'Japanese' }],
  });

  assert.deepEqual(stream, {
    url: 'http://127.0.0.1:8080/video/master-token',
    quality: '1080p',
    headers: { Referer: 'https://origin.example/' },
    subtitles: [{ url: 'http://127.0.0.1:8080/video/sub-token', lang: 'English' }],
    audios: [{ url: 'http://127.0.0.1:8080/video/audio-token', lang: 'Japanese' }],
  });
});

test('resolveStream returns null when the extension resolved no media url', () => {
  assert.equal(resolveStream({ url: 'https://origin.example/embed/1', quality: '1080p' }), null);
  assert.equal(resolveStream({ videoUrl: '' }), null);
});

test('resolveStream drops tracks without a url and defaults a missing lang', () => {
  const stream = resolveStream({
    videoUrl: 'http://127.0.0.1:8080/video/master-token',
    subtitleTracks: [{ lang: 'English' }, { url: 'http://127.0.0.1:8080/video/sub-token' }],
  });
  assert.deepEqual(stream?.subtitles, [{ url: 'http://127.0.0.1:8080/video/sub-token', lang: '' }]);
  assert.equal(stream?.quality, '');
});
