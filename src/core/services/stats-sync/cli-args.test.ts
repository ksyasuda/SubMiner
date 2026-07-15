import test from 'node:test';
import assert from 'node:assert/strict';
import { extractSyncCliTokens, parseSyncCliTokens } from './cli-args';

test('extractSyncCliTokens returns tokens after --sync-cli', () => {
  assert.equal(extractSyncCliTokens(['/bin/electron', '/app']), null);
  assert.deepEqual(extractSyncCliTokens(['/bin/electron', '/app', '--sync-cli', 'sync', 'host']), [
    'sync',
    'host',
  ]);
  // A repeated flag (e.g. resolved remote command + forwarded argv) is ignored.
  assert.deepEqual(extractSyncCliTokens(['/app', '--sync-cli', '--sync-cli', 'sync', 'h']), [
    'sync',
    'h',
  ]);
});

test('parseSyncCliTokens handles help, version, and run modes', () => {
  assert.deepEqual(parseSyncCliTokens(['--help']), { kind: 'help' });
  assert.deepEqual(parseSyncCliTokens(['--version']), { kind: 'version' });

  const run = parseSyncCliTokens(['sync', 'media-box', '--pull', '--force', '--json']);
  assert.equal(run.kind, 'run');
  if (run.kind === 'run') {
    assert.equal(run.args.syncHost, 'media-box');
    assert.equal(run.args.syncDirection, 'pull');
    assert.equal(run.args.syncForce, true);
    assert.equal(run.args.syncJson, true);
  }

  const snapshot = parseSyncCliTokens(['sync', '--snapshot', '/tmp/x.sqlite', '--db', '/tmp/db']);
  assert.equal(snapshot.kind, 'run');
  if (snapshot.kind === 'run') {
    assert.equal(snapshot.args.syncSnapshotPath, '/tmp/x.sqlite');
    assert.equal(snapshot.args.syncDbPath, '/tmp/db');
  }
});

test('parseSyncCliTokens handles the temp-dir protocol modes', () => {
  const makeTemp = parseSyncCliTokens(['sync', '--make-temp']);
  assert.equal(makeTemp.kind, 'run');
  if (makeTemp.kind === 'run') assert.equal(makeTemp.args.syncMakeTemp, true);

  const removeTemp = parseSyncCliTokens(['sync', '--remove-temp', '/tmp/subminer-sync-x']);
  assert.equal(removeTemp.kind, 'run');
  if (removeTemp.kind === 'run') {
    assert.equal(removeTemp.args.syncRemoveTempPath, '/tmp/subminer-sync-x');
  }

  assert.equal(parseSyncCliTokens(['sync', '--make-temp', 'host']).kind, 'error');
  assert.equal(
    parseSyncCliTokens(['sync', '--make-temp', '--remove-temp', '/tmp/x']).kind,
    'error',
  );
});

test('parseSyncCliTokens owns the sync CLI validation rules', () => {
  assert.equal(parseSyncCliTokens([]).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', 'h', '--push', '--pull']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', '--snapshot', '/tmp/x', '--push']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', '--check']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', '--check', '--snapshot', '/tmp/x', 'h']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', 'h', '--snapshot', '/tmp/x']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', 'h', '--bogus']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', 'h', 'extra']).kind, 'error');
  assert.equal(parseSyncCliTokens(['sync', '--snapshot']).kind, 'error');
});

test('parseSyncCliTokens rejects an option-like token where a value is required', () => {
  // Without this guard `--snapshot --force` writes a snapshot to a file literally
  // named "--force" and silently drops the flag.
  assert.deepEqual(parseSyncCliTokens(['sync', '--snapshot', '--force']), {
    kind: 'error',
    message: 'Missing value for --snapshot.',
  });
  assert.deepEqual(parseSyncCliTokens(['sync', '--remove-temp', '--force']), {
    kind: 'error',
    message: 'Missing value for --remove-temp.',
  });
  // The `--flag=<value>` form still accepts values that begin with "-".
  const parsed = parseSyncCliTokens(['sync', '--snapshot=-weird-name.sqlite']);
  assert.equal(parsed.kind, 'run');
  assert.equal(parsed.kind === 'run' && parsed.args.syncSnapshotPath, '-weird-name.sqlite');
});
