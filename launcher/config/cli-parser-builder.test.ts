import assert from 'node:assert/strict';
import test from 'node:test';
import { parseCliPrograms, resolveTopLevelCommand } from './cli-parser-builder.js';
import { SUBTITLE_GENERATION_MODELS } from '../../src/shared/subtitle-generation-model-catalog.js';

test('generate-subs accepts all downloadable multilingual model variants', () => {
  for (const { id } of SUBTITLE_GENERATION_MODELS) {
    const { invocations } = parseCliPrograms(['generate-subs', '--model', id], 'subminer');
    assert.equal(invocations.generateSubtitles?.managedModel, id);
  }
});

test('generate-subs parses local generation options separately from YouTube options', () => {
  const result = parseCliPrograms(
    [
      'generate-subs',
      'episode.mkv',
      '--download-model',
      '--model',
      'medium',
      '--output',
      'episode.ja.srt',
      '--audio-stream',
      '2',
    ],
    'subminer',
  );
  assert.deepEqual(result.invocations.generateSubtitles, {
    mediaPath: 'episode.mkv',
    downloadModel: true,
    managedModel: 'medium',
    modelPath: undefined,
    outputPath: 'episode.ja.srt',
    audioStreamIndex: 2,
  });
  assert.equal(
    parseCliPrograms(['generate-subs'], 'subminer').invocations.generateSubtitles?.mediaPath,
    undefined,
  );
  assert.equal(
    parseCliPrograms(['generate-subs', '--model-path', '/models/ggml.bin'], 'subminer').invocations
      .generateSubtitles?.modelPath,
    '/models/ggml.bin',
  );
});

test('generate-subs rejects conflicting models and malformed audio stream indices', () => {
  for (const flags of [
    ['--model', 'tiny.en'],
    ['--model', 'small.en-q5_1'],
    ['--model', 'toString'],
    ['--audio-stream', '-1'],
    ['--audio-stream', '1.5'],
    ['--model-path', '/model.bin', '--download-model'],
    ['--model-path', '/model.bin', '--model', 'small'],
  ])
    assert.throws(() => parseCliPrograms(['generate-subs', ...flags], 'subminer'), /Generation/);
});

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

  const removeTemp = parseCliPrograms(
    ['sync', '--remove-temp', '/tmp/subminer-sync-x'],
    'subminer',
  );
  assert.deepEqual(removeTemp.invocations.syncCliTokens, ['--remove-temp', '/tmp/subminer-sync-x']);
});

test('parseCliPrograms forwards transfer cache keys with both sync temp helpers', () => {
  const key = 'a'.repeat(64);
  for (const helper of [['--make-temp'], ['--remove-temp', '/tmp/subminer-sync-x']]) {
    const tokens = [...helper, '--transfer-cache', key];
    const result = parseCliPrograms(['sync', ...tokens], 'subminer');
    assert.equal(result.invocations.syncTriggered, true);
    assert.deepEqual(result.invocations.syncCliTokens, tokens);
  }
});

test('parseCliPrograms rejects sync --ui with --transfer-cache', () => {
  assert.throws(
    () => parseCliPrograms(['sync', '--ui', '--transfer-cache', 'a'.repeat(64)], 'subminer'),
    { message: 'Sync --ui cannot be combined with other sync options.' },
  );
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

test('parseCliPrograms rejects sync --ui with --remote-cmd', () => {
  assert.throws(
    () => parseCliPrograms(['sync', '--ui', '--remote-cmd', '/opt/SubMiner.AppImage'], 'subminer'),
    { message: 'Sync --ui cannot be combined with other sync options.' },
  );
});

test('parseCliPrograms rejects sync --ui with --db', () => {
  assert.throws(() => parseCliPrograms(['sync', '--ui', '--db', '/tmp/db.sqlite'], 'subminer'), {
    message: 'Sync --ui cannot be combined with other sync options.',
  });
});
