import test from 'node:test';
import assert from 'node:assert/strict';
import { codecToExtension, getSubsyncConfig, resolveLocalMediaPath } from './utils';

test('codecToExtension maps stream/web formats to ffmpeg extractable extensions', () => {
  assert.equal(codecToExtension('subrip'), 'srt');
  assert.equal(codecToExtension('webvtt'), 'vtt');
  assert.equal(codecToExtension('vtt'), 'vtt');
  assert.equal(codecToExtension('ttml'), 'ttml');
});

test('codecToExtension returns null for unsupported codecs', () => {
  assert.equal(codecToExtension('unsupported-codec'), null);
});

test('getSubsyncConfig defaults replace to true', () => {
  assert.equal(getSubsyncConfig(undefined).replace, true);
  assert.equal(getSubsyncConfig({}).replace, true);
});

test('getSubsyncConfig respects explicit replace value', () => {
  assert.equal(getSubsyncConfig({ replace: false }).replace, false);
  assert.equal(getSubsyncConfig({ replace: true }).replace, true);
});

test('resolveLocalMediaPath decodes file URLs mpv reports for dropped files', () => {
  assert.equal(
    resolveLocalMediaPath('file:///home/user/subs/%E3%83%AF%E3%83%B3%20sdh.srt'),
    '/home/user/subs/ワン sdh.srt',
  );
  assert.equal(resolveLocalMediaPath('FILE:///tmp/ref.srt'), '/tmp/ref.srt');
});

test('resolveLocalMediaPath leaves plain paths and stream URLs alone', () => {
  assert.equal(resolveLocalMediaPath('/tmp/ref.srt'), '/tmp/ref.srt');
  assert.equal(
    resolveLocalMediaPath('https://jellyfin.example/subs/eng.srt'),
    'https://jellyfin.example/subs/eng.srt',
  );
  // A UNC host has no local path; the caller's own error is the useful one.
  assert.equal(resolveLocalMediaPath('file://host/share/ref.srt'), 'file://host/share/ref.srt');
});
