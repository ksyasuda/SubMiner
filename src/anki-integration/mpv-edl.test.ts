import assert from 'node:assert/strict';
import test from 'node:test';

import { extractFileUrlsFromMpvEdlSource } from './mpv-edl';
import { toMpvEdlValue } from './mpv-edl-test-utils';

test('extractFileUrlsFromMpvEdlSource honors length-guarded file values', () => {
  const url =
    'https://rr1---sn.example.googlevideo.com/videoplayback?mime=video%2Fmp4&mn=sn-a,sn-b&lsig=abc%3D';
  const source = `edl://!new_stream;${toMpvEdlValue(url)},title=clip,length=73`;

  assert.deepEqual(extractFileUrlsFromMpvEdlSource(source), [url]);
});

test('extractFileUrlsFromMpvEdlSource reads file parameters', () => {
  const initUrl = 'https://init.example/init.mp4';
  const fileUrl = 'https://video.example/videoplayback?mime=video%2Fmp4';
  const source = `edl://!mp4_dash,init=${toMpvEdlValue(initUrl)};file=${toMpvEdlValue(
    fileUrl,
  )},length=42`;

  assert.deepEqual(extractFileUrlsFromMpvEdlSource(source), [fileUrl]);
});

test('extractFileUrlsFromMpvEdlSource aggregates file URLs across entries', () => {
  const audioUrl = 'https://audio.example/videoplayback?mime=audio%2Fwebm';
  const videoUrl = 'https://video.example/videoplayback?mime=video%2Fmp4';
  const source = [
    'edl://!new_stream',
    toMpvEdlValue(audioUrl),
    '!new_stream',
    `file=${toMpvEdlValue(videoUrl)},length=50`,
    '!global_tags,title=test',
  ].join(';');

  assert.deepEqual(extractFileUrlsFromMpvEdlSource(source), [audioUrl, videoUrl]);
});
