import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { Args } from '../types.js';
import { createEmptyMergeSummary } from '../sync/sync-shared.js';
import type { LauncherCommandContext } from './context.js';
import { runSyncCommand, type SyncCommandDeps } from './sync-command.js';

function makeContext(overrides: Partial<Args>): LauncherCommandContext {
  return {
    args: {
      sync: true,
      syncHost: '',
      syncSnapshotPath: '',
      syncMergePath: '',
      syncRemoteCmd: '',
      syncDbPath: '',
      syncForce: false,
      logLevel: 'warn',
      ...overrides,
    } as Args,
    scriptPath: '/tmp/subminer',
    scriptName: 'subminer',
    mpvSocketPath: '',
    pluginRuntimeConfig: {},
    appPath: null,
    launcherJellyfinConfig: {},
    processAdapter: process,
  } as unknown as LauncherCommandContext;
}

function ok(stdout = ''): { status: number; stdout: string; stderr: string } {
  return { status: 0, stdout, stderr: '' };
}

test('runSyncCommand dispatches snapshot, merge, host, and missing-target modes', () => {
  const calls: string[] = [];
  const deps: Partial<SyncCommandDeps> = {
    createDbSnapshot: (dbPath: string, outPath: string) => {
      calls.push(`snapshot:${dbPath}->${outPath}`);
    },
    mergeSnapshotIntoDb: (dbPath: string, snapshotPath: string) => {
      calls.push(`merge:${dbPath}<-${snapshotPath}`);
      return createEmptyMergeSummary();
    },
    formatMergeSummary: () => 'summary',
    ensureTrackerQuiescent: () => {
      calls.push('quiescent');
    },
    assertSafeSshHost: (host: string) => {
      calls.push(`host:${host}`);
    },
    resolveRemoteSubminerCommand: () => 'subminer',
    runSsh: (_host: string, command: string) => {
      calls.push(`ssh:${command}`);
      return command.startsWith('mktemp ') ? ok('/tmp/subminer-sync.remote\n') : ok();
    },
    runScp: (from: string, to: string) => {
      calls.push(`scp:${from}->${to}`);
    },
    fail: (message: string): never => {
      throw new Error(message);
    },
  };

  assert.equal(
    runSyncCommand(
      makeContext({ syncDbPath: '/tmp/local.sqlite', syncSnapshotPath: '/tmp/out.sqlite' }),
      deps,
    ),
    true,
  );
  assert.ok(calls.includes('snapshot:/tmp/local.sqlite->/tmp/out.sqlite'));

  runSyncCommand(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncMergePath: '/tmp/in.sqlite' }),
    deps,
  );
  assert.ok(calls.includes('quiescent'));
  assert.ok(calls.includes('merge:/tmp/local.sqlite<-/tmp/in.sqlite'));

  runSyncCommand(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }), deps);
  assert.ok(calls.includes('host:media-box'));

  assert.throws(
    () => runSyncCommand(makeContext({ syncDbPath: '/tmp/local.sqlite' }), deps),
    /sync requires a host, --snapshot <file>, or --merge <file>/,
  );
});

test('runHostSync keeps tracker quiescent through local and remote merge and cleans up after failure', () => {
  const calls: string[] = [];
  let localTmpDir = '';
  const deps: Partial<SyncCommandDeps> = {
    createDbSnapshot: (_dbPath: string, outPath: string) => {
      calls.push(`snapshot:${outPath}`);
      fs.writeFileSync(outPath, 'snapshot');
    },
    mergeSnapshotIntoDb: () => {
      calls.push('local-merge');
      return createEmptyMergeSummary();
    },
    formatMergeSummary: () => 'summary',
    ensureTrackerQuiescent: () => {
      calls.push('quiescent');
    },
    assertSafeSshHost: () => {},
    resolveRemoteSubminerCommand: () => 'subminer',
    mkdtempSync: ((prefix: string) => {
      localTmpDir = fs.mkdtempSync(path.join(os.tmpdir(), path.basename(prefix)));
      return localTmpDir;
    }) as typeof fs.mkdtempSync,
    runSsh: (_host: string, command: string) => {
      calls.push(`ssh:${command}`);
      if (command.startsWith('mktemp ')) return ok('/tmp/subminer-sync.remote\n');
      if (command.includes(' sync --snapshot ')) return ok();
      if (command.includes(' sync --merge ')) {
        return { status: 9, stdout: 'remote output', stderr: 'remote merge exploded' };
      }
      return ok();
    },
    runScp: (from: string, to: string) => {
      calls.push(`scp:${from}->${to}`);
      if (!to.includes(':')) fs.writeFileSync(to, 'pulled');
    },
  };

  assert.throws(
    () =>
      runSyncCommand(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }), deps),
    /Remote merge failed on media-box[\s\S]*remote merge exploded/,
  );
  assert.equal(calls.filter((call) => call === 'quiescent').length, 3);
  assert.ok(calls.indexOf('quiescent') < calls.findIndex((call) => call.startsWith('snapshot:')));
  assert.ok(calls.includes('local-merge'));
  assert.ok(calls.some((call) => call.startsWith('ssh:rm -rf ')));
  assert.equal(fs.existsSync(localTmpDir), false);
});

test('runHostSync includes remote snapshot stderr in failures', () => {
  const deps: Partial<SyncCommandDeps> = {
    createDbSnapshot: (_dbPath: string, outPath: string) => {
      fs.writeFileSync(outPath, 'snapshot');
    },
    ensureTrackerQuiescent: () => {},
    assertSafeSshHost: () => {},
    resolveRemoteSubminerCommand: () => 'subminer',
    runSsh: (_host: string, command: string) => {
      if (command.startsWith('mktemp ')) return ok('/tmp/subminer-sync.remote\n');
      if (command.includes(' sync --snapshot ')) {
        return { status: 5, stdout: '', stderr: 'snapshot permission denied' };
      }
      return ok();
    },
    runScp: () => {},
  };

  assert.throws(
    () =>
      runSyncCommand(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }), deps),
    /Remote snapshot failed on media-box[\s\S]*snapshot permission denied/,
  );
});
