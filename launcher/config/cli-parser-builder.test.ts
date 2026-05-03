import assert from 'node:assert/strict';
import test from 'node:test';
import { parseCliPrograms, resolveTopLevelCommand } from './cli-parser-builder.js';

test('resolveTopLevelCommand skips root options and finds the first command', () => {
  assert.deepEqual(resolveTopLevelCommand(['--backend', 'macos', 'config', 'show']), {
    name: 'config',
    index: 2,
  });
});

test('resolveTopLevelCommand respects the app alias after root options', () => {
  assert.deepEqual(resolveTopLevelCommand(['--log-level', 'debug', 'bin', '--foo']), {
    name: 'bin',
    index: 2,
  });
});

test('parseCliPrograms keeps root options and target when no command is present', () => {
  const result = parseCliPrograms(['--backend', 'windows', '/tmp/movie.mkv'], 'subminer');

  assert.equal(result.options.backend, 'windows');
  assert.equal(result.rootTarget, '/tmp/movie.mkv');
  assert.equal(result.invocations.appInvocation, null);
});

test('parseCliPrograms routes app alias arguments through passthrough mode', () => {
  const result = parseCliPrograms(
    ['--backend', 'windows', 'bin', '--anilist', '--log-level', 'debug'],
    'subminer',
  );

  assert.equal(result.options.backend, 'windows');
  assert.deepEqual(result.invocations.appInvocation, {
    appArgs: ['--anilist', '--log-level', 'debug'],
  });
});

test('parseCliPrograms captures texthooker browser-open flag', () => {
  const result = parseCliPrograms(['texthooker', '-o'], 'subminer');

  assert.equal(result.invocations.texthookerTriggered, true);
  assert.equal(result.invocations.texthookerOpenBrowser, true);
});
