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

test('parseCliPrograms lowers sync options into app-owned CLI tokens', () => {
  const push = parseCliPrograms(['sync', 'media-box', '--push'], 'subminer');
  assert.equal(push.invocations.syncTriggered, true);
  assert.deepEqual(push.invocations.syncCliTokens, ['media-box', '--push']);

  const pull = parseCliPrograms(['sync', 'media-box', '--pull'], 'subminer');
  assert.deepEqual(pull.invocations.syncCliTokens, ['media-box', '--pull']);

  const check = parseCliPrograms(['sync', 'media-box', '--check', '--json'], 'subminer');
  assert.deepEqual(check.invocations.syncCliTokens, ['media-box', '--check', '--json']);

  const full = parseCliPrograms(
    [
      'sync',
      'media-box',
      '--remote-cmd',
      '/opt/SubMiner.AppImage',
      '--db',
      '/tmp/db.sqlite',
      '--force',
      '--log-level',
      'debug',
    ],
    'subminer',
  );
  assert.deepEqual(full.invocations.syncCliTokens, [
    'media-box',
    '--remote-cmd',
    '/opt/SubMiner.AppImage',
    '--db',
    '/tmp/db.sqlite',
    '--force',
  ]);
  assert.equal(full.invocations.syncLogLevel, 'debug');

  const snapshot = parseCliPrograms(['sync', '--snapshot', '/tmp/out.sqlite'], 'subminer');
  assert.deepEqual(snapshot.invocations.syncCliTokens, ['--snapshot', '/tmp/out.sqlite']);

  const merge = parseCliPrograms(['sync', '--merge', '/tmp/in.sqlite'], 'subminer');
  assert.deepEqual(merge.invocations.syncCliTokens, ['--merge', '/tmp/in.sqlite']);

  const makeTemp = parseCliPrograms(['sync', '--make-temp'], 'subminer');
  assert.deepEqual(makeTemp.invocations.syncCliTokens, ['--make-temp']);

  const removeTemp = parseCliPrograms(['sync', '--remove-temp', '/tmp/subminer-sync-x'], 'subminer');
  assert.deepEqual(removeTemp.invocations.syncCliTokens, ['--remove-temp', '/tmp/subminer-sync-x']);
});

test('parseCliPrograms leaves sync validation to the app parser', () => {
  // Invalid combinations are forwarded; the app's parseSyncCliTokens rejects them.
  const invalid = parseCliPrograms(['sync', 'media-box', '--push', '--pull'], 'subminer');
  assert.equal(invalid.invocations.syncTriggered, true);
  assert.deepEqual(invalid.invocations.syncCliTokens, ['media-box', '--push', '--pull']);

  const empty = parseCliPrograms(['sync'], 'subminer');
  assert.equal(empty.invocations.syncTriggered, true);
  assert.deepEqual(empty.invocations.syncCliTokens, []);
});

test('parseCliPrograms captures sync --ui', () => {
  const result = parseCliPrograms(['sync', '--ui'], 'subminer');
  assert.equal(result.invocations.syncUiTriggered, true);
  assert.equal(result.invocations.syncTriggered, false);

  assert.throws(
    () => parseCliPrograms(['sync', 'media-box', '--ui'], 'subminer'),
    /--ui cannot be combined/,
  );
});
