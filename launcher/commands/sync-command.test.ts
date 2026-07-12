import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { Args } from '../types.js';
import { createEmptyMergeSummary } from '../sync/sync-shared.js';
import type { LauncherCommandContext } from './context.js';
import { ensureTrackerQuiescent, runSyncCommand, type SyncCommandDeps } from './sync-command.js';

function makeContext(overrides: Partial<Args>): LauncherCommandContext {
  return {
    args: {
      sync: true,
      syncHost: '',
      syncSnapshotPath: '',
      syncMergePath: '',
      syncDirection: 'both',
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

// Every host-sync test must stub recordHostSyncResult: the default writes the
// real sync-hosts.json in the user's config directory.
const noRecord: Pick<SyncCommandDeps, 'recordHostSyncResult'> = {
  recordHostSyncResult: () => {},
};

test('ensureTrackerQuiescent ignores stale sockets but rejects live sockets', async () => {
  const context = makeContext({ syncDbPath: '/tmp/local.sqlite' });
  context.mpvSocketPath = '/tmp/subminer-socket';
  let socketConnectable = false;
  const deps: Partial<SyncCommandDeps> = {
    realpathSync: (() => '/tracker.sqlite') as unknown as typeof fs.realpathSync,
    findLiveStatsDaemonPid: () => null,
    canConnectUnixSocket: async () => socketConnectable,
    fail: (message: string): never => {
      throw new Error(message);
    },
  };

  await ensureTrackerQuiescent(context, '/tmp/local.sqlite', deps);

  socketConnectable = true;
  await assert.rejects(
    async () => ensureTrackerQuiescent(context, '/tmp/local.sqlite', deps),
    /mpv\/SubMiner session appears to be running/,
  );
});

test('runSyncCommand dispatches snapshot, merge, host, and missing-target modes', async () => {
  const calls: string[] = [];
  const deps: Partial<SyncCommandDeps> = {
    ...noRecord,
    createDbSnapshot: (dbPath: string, outPath: string) => {
      calls.push(`snapshot:${dbPath}->${outPath}`);
    },
    mergeSnapshotIntoDb: (dbPath: string, snapshotPath: string) => {
      calls.push(`merge:${dbPath}<-${snapshotPath}`);
      return createEmptyMergeSummary();
    },
    formatMergeSummary: () => 'summary',
    ensureTrackerQuiescent: async () => {
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
    await runSyncCommand(
      makeContext({ syncDbPath: '/tmp/local.sqlite', syncSnapshotPath: '/tmp/out.sqlite' }),
      deps,
    ),
    true,
  );
  assert.ok(calls.includes('snapshot:/tmp/local.sqlite->/tmp/out.sqlite'));

  await runSyncCommand(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncMergePath: '/tmp/in.sqlite' }),
    deps,
  );
  assert.ok(calls.includes('quiescent'));
  assert.ok(calls.includes('merge:/tmp/local.sqlite<-/tmp/in.sqlite'));

  await runSyncCommand(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }),
    deps,
  );
  assert.ok(calls.includes('host:media-box'));

  await assert.rejects(
    () => runSyncCommand(makeContext({ syncDbPath: '/tmp/local.sqlite' }), deps),
    /sync requires a host, --snapshot <file>, or --merge <file>/,
  );
});

test('runHostSync keeps tracker quiescent through local and remote merge and cleans up after failure', async () => {
  const calls: string[] = [];
  let localTmpDir = '';
  const deps: Partial<SyncCommandDeps> = {
    ...noRecord,
    createDbSnapshot: (_dbPath: string, outPath: string) => {
      calls.push(`snapshot:${outPath}`);
      fs.writeFileSync(outPath, 'snapshot');
    },
    mergeSnapshotIntoDb: () => {
      calls.push('local-merge');
      return createEmptyMergeSummary();
    },
    formatMergeSummary: () => 'summary',
    ensureTrackerQuiescent: async () => {
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

  await assert.rejects(
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

test('runHostSync includes remote snapshot stderr in failures', async () => {
  const deps: Partial<SyncCommandDeps> = {
    ...noRecord,
    createDbSnapshot: (_dbPath: string, outPath: string) => {
      fs.writeFileSync(outPath, 'snapshot');
    },
    ensureTrackerQuiescent: async () => {},
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

  await assert.rejects(
    () =>
      runSyncCommand(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }), deps),
    /Remote snapshot failed on media-box[\s\S]*snapshot permission denied/,
  );
});

function makeDirectionDeps(calls: string[]): Partial<SyncCommandDeps> {
  return {
    ...noRecord,
    createDbSnapshot: (_dbPath: string, outPath: string) => {
      calls.push(`snapshot:${outPath}`);
      fs.writeFileSync(outPath, 'snapshot');
    },
    mergeSnapshotIntoDb: () => {
      calls.push('local-merge');
      return createEmptyMergeSummary();
    },
    formatMergeSummary: () => 'summary',
    ensureTrackerQuiescent: async () => {
      calls.push('quiescent');
    },
    assertSafeSshHost: () => {},
    resolveRemoteSubminerCommand: () => 'subminer',
    runSsh: (_host: string, command: string) => {
      calls.push(`ssh:${command}`);
      if (command.startsWith('mktemp ')) return ok('/tmp/subminer-sync.remote\n');
      return ok();
    },
    runScp: (from: string, to: string) => {
      calls.push(`scp:${from}->${to}`);
      if (!to.includes(':')) fs.writeFileSync(to, 'pulled');
    },
  };
}

test('runHostSync push only snapshots locally and merges remotely', async () => {
  const calls: string[] = [];

  await runSyncCommand(
    makeContext({
      syncDbPath: '/tmp/local.sqlite',
      syncHost: 'media-box',
      syncDirection: 'push',
    }),
    makeDirectionDeps(calls),
  );

  assert.ok(calls.some((call) => call.startsWith('snapshot:')));
  assert.ok(calls.some((call) => call.includes(' sync --merge ')));
  assert.ok(calls.some((call) => call.startsWith('scp:') && call.includes('->media-box:')));
  assert.ok(!calls.some((call) => call.includes(' sync --snapshot ')));
  assert.ok(!calls.includes('local-merge'));
});

test('runHostSync pull only snapshots remotely and merges locally', async () => {
  const calls: string[] = [];

  await runSyncCommand(
    makeContext({
      syncDbPath: '/tmp/local.sqlite',
      syncHost: 'media-box',
      syncDirection: 'pull',
    }),
    makeDirectionDeps(calls),
  );

  assert.ok(calls.some((call) => call.includes(' sync --snapshot ')));
  assert.ok(calls.some((call) => call.startsWith('scp:media-box:')));
  assert.ok(calls.includes('local-merge'));
  assert.ok(!calls.some((call) => call.startsWith('snapshot:')));
  assert.ok(!calls.some((call) => call.includes(' sync --merge ')));
});

test('runSyncCommand --json emits NDJSON progress events for a two-way host sync', async () => {
  const lines: string[] = [];
  const deps: Partial<SyncCommandDeps> = {
    ...makeDirectionDeps([]),
    consoleLog: (line: string) => {
      lines.push(line);
    },
    writeStdout: ((chunk: string) => {
      lines.push(`stdout:${chunk}`);
      return true;
    }) as unknown as typeof process.stdout.write,
    runSsh: (_host: string, command: string) => {
      if (command.startsWith('mktemp ')) return ok('/tmp/subminer-sync.remote\n');
      if (command.includes(' sync --merge ')) return ok('remote summary text\n');
      return ok();
    },
    recordHostSyncResult: () => {},
  };

  await runSyncCommand(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box', syncJson: true }),
    deps,
  );

  const events = lines.map((line) => JSON.parse(line));
  const types = events.map((event) => event.type);
  assert.ok(types.includes('stage'));
  assert.ok(events.some((event) => event.type === 'stage' && event.stage === 'snapshot-local'));
  assert.ok(events.some((event) => event.type === 'merge-summary' && event.target === 'local'));
  assert.ok(
    events.some(
      (event) => event.type === 'remote-output' && event.text.includes('remote summary text'),
    ),
  );
  const last = events[events.length - 1];
  assert.deepEqual(last, { type: 'result', ok: true, error: null });
  assert.ok(!lines.some((line) => line.startsWith('stdout:')));
});

test('runSyncCommand --json emits an error result when the sync fails', async () => {
  const lines: string[] = [];
  const deps: Partial<SyncCommandDeps> = {
    ...makeDirectionDeps([]),
    consoleLog: (line: string) => {
      lines.push(line);
    },
    writeStdout: (() => true) as unknown as typeof process.stdout.write,
    runSsh: (_host: string, command: string) => {
      if (command.startsWith('mktemp ')) return ok('/tmp/subminer-sync.remote\n');
      if (command.includes(' sync --merge ')) {
        return { status: 9, stdout: '', stderr: 'remote merge exploded' };
      }
      return ok();
    },
    recordHostSyncResult: () => {},
  };

  await assert.rejects(() =>
    runSyncCommand(
      makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box', syncJson: true }),
      deps,
    ),
  );

  const events = lines.map((line) => JSON.parse(line));
  const last = events[events.length - 1];
  assert.equal(last.type, 'result');
  assert.equal(last.ok, false);
  assert.match(last.error, /Remote merge failed/);
});

test('runSyncCommand records host sync results for saved-host bookkeeping', async () => {
  const recorded: Array<{ host: string; status: string; detail: string | null }> = [];
  const baseDeps = makeDirectionDeps([]);
  const deps: Partial<SyncCommandDeps> = {
    ...baseDeps,
    recordHostSyncResult: (host: string, status: 'success' | 'error', detail: string | null) => {
      recorded.push({ host, status, detail });
    },
  };

  await runSyncCommand(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }),
    deps,
  );
  assert.equal(recorded.length, 1);
  assert.equal(recorded[0]!.host, 'media-box');
  assert.equal(recorded[0]!.status, 'success');

  const failingDeps: Partial<SyncCommandDeps> = {
    ...deps,
    runSsh: (_host: string, command: string) => {
      if (command.startsWith('mktemp ')) return ok('/tmp/subminer-sync.remote\n');
      if (command.includes(' sync --merge ')) {
        return { status: 9, stdout: '', stderr: 'boom' };
      }
      return ok();
    },
  };
  await assert.rejects(() =>
    runSyncCommand(
      makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }),
      failingDeps,
    ),
  );
  assert.equal(recorded.length, 2);
  assert.equal(recorded[1]!.status, 'error');
});

test('runSyncCommand --check --json reports ssh and remote launcher status', async () => {
  const lines: string[] = [];
  const deps: Partial<SyncCommandDeps> = {
    ensureTrackerQuiescent: async () => {},
    assertSafeSshHost: () => {},
    resolveRemoteSubminerCommand: () => 'subminer',
    consoleLog: (line: string) => {
      lines.push(line);
    },
    runSsh: (_host: string, command: string) => {
      if (command.includes('--version')) return ok('SubMiner 0.18.0\n');
      return ok('subminer-check-ok\n');
    },
    fail: (message: string): never => {
      throw new Error(message);
    },
  };

  assert.equal(
    await runSyncCommand(
      makeContext({
        syncDbPath: '/tmp/local.sqlite',
        syncHost: 'media-box',
        syncCheck: true,
        syncJson: true,
      }),
      deps,
    ),
    true,
  );

  const events = lines.map((line) => JSON.parse(line));
  const check = events.find((event) => event.type === 'check-result');
  assert.ok(check);
  assert.equal(check.host, 'media-box');
  assert.equal(check.sshOk, true);
  assert.equal(check.remoteCommand, 'subminer');
  assert.equal(check.remoteVersion, 'SubMiner 0.18.0');
  assert.equal(check.ok, true);
  assert.equal(check.error, null);
});

test('runSyncCommand --check --json reports failures without throwing early', async () => {
  const lines: string[] = [];
  const deps: Partial<SyncCommandDeps> = {
    ensureTrackerQuiescent: async () => {},
    assertSafeSshHost: () => {},
    resolveRemoteSubminerCommand: () => {
      throw new Error('Could not find a runnable "subminer" on media-box.');
    },
    consoleLog: (line: string) => {
      lines.push(line);
    },
    runSsh: () => ok('subminer-check-ok\n'),
    fail: (message: string): never => {
      throw new Error(message);
    },
  };

  await assert.rejects(() =>
    runSyncCommand(
      makeContext({
        syncDbPath: '/tmp/local.sqlite',
        syncHost: 'media-box',
        syncCheck: true,
        syncJson: true,
      }),
      deps,
    ),
  );

  const events = lines.map((line) => JSON.parse(line));
  const check = events.find((event) => event.type === 'check-result');
  assert.ok(check);
  assert.equal(check.sshOk, true);
  assert.equal(check.ok, false);
  assert.match(check.error, /Could not find a runnable/);
});
