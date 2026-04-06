import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildMpvLaunchModeArgs,
  parseMpvLaunchMode,
  resolveRawMpvLaunchMode,
} from './mpv-launch-mode';

test('parseMpvLaunchMode normalizes valid string values', () => {
  assert.equal(parseMpvLaunchMode(' fullscreen '), 'fullscreen');
  assert.equal(parseMpvLaunchMode('MAXIMIZED'), 'maximized');
  assert.equal(parseMpvLaunchMode('normal'), 'normal');
});

test('resolveRawMpvLaunchMode falls back to deprecated fullscreen alias when launchMode is absent', () => {
  assert.equal(resolveRawMpvLaunchMode({ startFullscreen: true }), 'fullscreen');
  assert.equal(resolveRawMpvLaunchMode({ startFullscreen: false }), 'normal');
});

test('resolveRawMpvLaunchMode ignores deprecated alias when launchMode is present', () => {
  assert.equal(
    resolveRawMpvLaunchMode({ launchMode: 'maximized', startFullscreen: true }),
    'maximized',
  );
  assert.equal(
    resolveRawMpvLaunchMode({ launchMode: 'cinema', startFullscreen: true }),
    undefined,
  );
});

test('buildMpvLaunchModeArgs returns the expected mpv flags', () => {
  assert.deepEqual(buildMpvLaunchModeArgs('normal'), []);
  assert.deepEqual(buildMpvLaunchModeArgs('maximized'), ['--window-maximized=yes']);
  assert.deepEqual(buildMpvLaunchModeArgs('fullscreen'), ['--fullscreen']);
});
