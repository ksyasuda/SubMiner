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
