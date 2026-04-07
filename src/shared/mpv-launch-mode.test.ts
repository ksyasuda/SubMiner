import assert from 'node:assert/strict';
import test from 'node:test';
import { buildMpvLaunchModeArgs, parseMpvLaunchMode } from './mpv-launch-mode';

test('parseMpvLaunchMode normalizes valid string values', () => {
  assert.equal(parseMpvLaunchMode(' fullscreen '), 'fullscreen');
  assert.equal(parseMpvLaunchMode('MAXIMIZED'), 'maximized');
  assert.equal(parseMpvLaunchMode('normal'), 'normal');
});

test('buildMpvLaunchModeArgs returns the expected mpv flags', () => {
  assert.deepEqual(buildMpvLaunchModeArgs('normal'), []);
  assert.deepEqual(buildMpvLaunchModeArgs('maximized'), ['--window-maximized=yes']);
  assert.deepEqual(buildMpvLaunchModeArgs('fullscreen'), ['--fullscreen']);
});
