import test from 'node:test';
import assert from 'node:assert/strict';
import { codecToExtension, getSubsyncConfig } from './utils';

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
