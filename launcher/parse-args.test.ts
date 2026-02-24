import test from 'node:test';
import assert from 'node:assert/strict';
import { parseArgs } from './config';

test('parseArgs captures passthrough args for app subcommand', () => {
  const parsed = parseArgs(['app', '--anilist', '--log-level', 'debug'], 'subminer', {});

  assert.equal(parsed.appPassthrough, true);
  assert.deepEqual(parsed.appArgs, ['--anilist', '--log-level', 'debug']);
});

test('parseArgs supports bin alias for app subcommand', () => {
  const parsed = parseArgs(['bin', '--anilist-status'], 'subminer', {});

  assert.equal(parsed.appPassthrough, true);
  assert.deepEqual(parsed.appArgs, ['--anilist-status']);
});

test('parseArgs keeps all args after app verbatim', () => {
  const parsed = parseArgs(['app', '--start', '--anilist-setup', '-h'], 'subminer', {});

  assert.equal(parsed.appPassthrough, true);
  assert.deepEqual(parsed.appArgs, ['--start', '--anilist-setup', '-h']);
});

test('parseArgs maps jellyfin play action and log-level override', () => {
  const parsed = parseArgs(['jellyfin', 'play', '--log-level', 'debug'], 'subminer', {});

  assert.equal(parsed.jellyfinPlay, true);
  assert.equal(parsed.logLevel, 'debug');
});

test('parseArgs forwards jellyfin password-store option', () => {
  const parsed = parseArgs(
    ['jf', 'setup', '--password-store', 'gnome-libsecret'],
    'subminer',
    {},
  );

  assert.equal(parsed.jellyfin, true);
  assert.equal(parsed.passwordStore, 'gnome-libsecret');
});

test('parseArgs maps config show action', () => {
  const parsed = parseArgs(['config', 'show'], 'subminer', {});

  assert.equal(parsed.configShow, true);
  assert.equal(parsed.configPath, false);
});

test('parseArgs maps mpv idle action', () => {
  const parsed = parseArgs(['mpv', 'idle'], 'subminer', {});

  assert.equal(parsed.mpvIdle, true);
  assert.equal(parsed.mpvStatus, false);
});
