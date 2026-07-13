import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createEmptyMergeSummary } from './shared';
import {
  ensureTrackerQuiescentFlow,
  runSyncFlow,
  type SyncFlowContext,
  type SyncFlowDeps,
} from './sync-flow';

function makeContext(overrides: Partial<SyncFlowContext['args']> = {}): SyncFlowContext {
  return {
    args: {
      syncHost: '',
      syncSnapshotPath: '',
      syncMergePath: '',
      syncDirection: 'both',
      syncRemoteCmd: '',
      syncDbPath: '',
      syncForce: false,
      syncJson: false,
      syncCheck: false,
      syncMakeTemp: false,
      syncRemoveTempPath: '',
      logLevel: 'warn',
      ...overrides,
    },
    mpvSocketPath: '',
  };
}

function ok(stdout = ''): { status: number; stdout: string; stderr: string } {
  return { status: 0, stdout, stderr: '' };
}

// recordHostSyncResult defaults to a no-op here: the real disk-writing
// implementations are bound by the entry points, never by the flow itself.
function makeDeps(overrides: Partial<SyncFlowDeps> = {}): SyncFlowDeps {
  return {
    createDbSnapshot: () => {},
    mergeSnapshotIntoDb: () => createEmptyMergeSummary(),
    findLiveStatsDaemonPid: () => null,
    assertSafeSshHost: () => {},
    detectRemoteShellFlavor: () => 'posix',
    resolveRemoteSubminerCommand: () => 'subminer',
    runScp: () => {},
    runSsh: () => ok(),
    canConnectUnixSocket: async () => false,
    realpathSync: (candidate) => candidate,
    mkdtempSync: (prefix) => fs.mkdtempSync(prefix),
    rmSync: (target, options) => fs.rmSync(target, options),
    consoleLog: () => {},
    writeStdout: () => true,
    ensureTrackerQuiescent: async () => {},
    emitEvent: () => {},
    recordHostSyncResult: () => {},
    resolveDefaultDbPath: () => '/tracker.sqlite',
    ...overrides,
  };
}

test('ensureTrackerQuiescentFlow ignores stale sockets but rejects live sockets', async () => {
  const context = makeContext({ syncDbPath: '/tmp/local.sqlite' });
  context.mpvSocketPath = '/tmp/subminer-socket';
  let socketConnectable = false;
  const deps = makeDeps({
    realpathSync: () => '/tracker.sqlite',
    canConnectUnixSocket: async () => socketConnectable,
  });

  await ensureTrackerQuiescentFlow(context, '/tmp/local.sqlite', deps);

  socketConnectable = true;
  await assert.rejects(
    async () => ensureTrackerQuiescentFlow(context, '/tmp/local.sqlite', deps),
    /mpv\/SubMiner session appears to be running/,
  );
});

test('ensureTrackerQuiescentFlow rejects a live stats daemon and honors --force', async () => {
  const context = makeContext({ syncDbPath: '/tmp/local.sqlite' });
  const deps = makeDeps({
    realpathSync: () => '/tracker.sqlite',
    findLiveStatsDaemonPid: () => 4242,
  });

  await assert.rejects(
    async () => ensureTrackerQuiescentFlow(context, '/tmp/local.sqlite', deps),
    /stats server is running \(pid 4242\)/,
  );

  const forced = makeContext({ syncDbPath: '/tmp/local.sqlite', syncForce: true });
  await ensureTrackerQuiescentFlow(forced, '/tmp/local.sqlite', deps);
});

test('runSyncFlow dispatches snapshot, merge, host, and missing-target modes', async () => {
  const calls: string[] = [];
  const deps = makeDeps({
    createDbSnapshot: (dbPath, outPath) => {
      calls.push(`snapshot:${dbPath}->${outPath}`);
    },
    mergeSnapshotIntoDb: (dbPath, snapshotPath) => {
      calls.push(`merge:${dbPath}<-${snapshotPath}`);
      return createEmptyMergeSummary();
    },
    ensureTrackerQuiescent: async () => {
      calls.push('quiescent');
    },
    assertSafeSshHost: (host) => {
      calls.push(`host:${host}`);
    },
    runSsh: (_host, command) => {
      calls.push(`ssh:${command}`);
      return command.includes(' sync --make-temp') ? ok('/tmp/subminer-sync-remote\n') : ok();
    },
    runScp: (from, to) => {
      calls.push(`scp:${from}->${to}`);
    },
  });

  await runSyncFlow(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncSnapshotPath: '/tmp/out.sqlite' }),
    deps,
  );
  assert.ok(calls.includes('snapshot:/tmp/local.sqlite->/tmp/out.sqlite'));
  assert.ok(
    calls.indexOf('quiescent') < calls.indexOf('snapshot:/tmp/local.sqlite->/tmp/out.sqlite'),
  );

  await runSyncFlow(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncMergePath: '/tmp/in.sqlite' }),
    deps,
  );
  assert.ok(calls.includes('quiescent'));
  assert.ok(calls.includes('merge:/tmp/local.sqlite<-/tmp/in.sqlite'));

  await runSyncFlow(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }), deps);
  assert.ok(calls.includes('host:media-box'));

  await assert.rejects(
    () => runSyncFlow(makeContext({ syncDbPath: '/tmp/local.sqlite' }), deps),
    /sync requires a host, --snapshot <file>, or --merge <file>/,
  );
});

function makeHostDeps(calls: string[], overrides: Partial<SyncFlowDeps> = {}): SyncFlowDeps {
  return makeDeps({
    createDbSnapshot: (_dbPath, outPath) => {
      calls.push(`snapshot:${outPath}`);
      fs.writeFileSync(outPath, 'snapshot');
    },
    mergeSnapshotIntoDb: () => {
      calls.push('local-merge');
      return createEmptyMergeSummary();
    },
    ensureTrackerQuiescent: async () => {
      calls.push('quiescent');
    },
    runSsh: (_host, command) => {
      calls.push(`ssh:${command}`);
      if (command.includes(' sync --make-temp')) return ok('/tmp/subminer-sync-remote\n');
      return ok();
    },
    runScp: (from, to) => {
      calls.push(`scp:${from}->${to}`);
      if (!to.includes(':')) fs.writeFileSync(to, 'pulled');
    },
    ...overrides,
  });
}

test('runHostSync keeps tracker quiescent through both merges and cleans up after failure', async () => {
  const calls: string[] = [];
  let localTmpDir = '';
  const deps = makeHostDeps(calls, {
    mkdtempSync: (prefix) => {
      localTmpDir = fs.mkdtempSync(prefix);
      return localTmpDir;
    },
    runSsh: (_host, command) => {
      calls.push(`ssh:${command}`);
      if (command.includes(' sync --make-temp')) return ok('/tmp/subminer-sync-remote\n');
      if (command.includes(' sync --merge ')) {
        return { status: 9, stdout: 'remote output', stderr: 'remote merge exploded' };
      }
      return ok();
    },
  });

  await assert.rejects(
    () =>
      runSyncFlow(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }), deps),
    /Remote merge failed on media-box[\s\S]*remote merge exploded/,
  );
  assert.equal(calls.filter((call) => call === 'quiescent').length, 3);
  assert.ok(calls.indexOf('quiescent') < calls.findIndex((call) => call.startsWith('snapshot:')));
  assert.ok(calls.includes('local-merge'));
  assert.ok(calls.some((call) => call.includes(' sync --remove-temp ')));
  assert.equal(fs.existsSync(localTmpDir), false);
});

test('runHostSync includes remote snapshot stderr in failures', async () => {
  const deps = makeHostDeps([], {
    runSsh: (_host, command) => {
      if (command.includes(' sync --make-temp')) return ok('/tmp/subminer-sync-remote\n');
      if (command.includes(' sync --snapshot ')) {
        return { status: 5, stdout: '', stderr: 'snapshot permission denied' };
      }
      return ok();
    },
  });

  await assert.rejects(
    () =>
      runSyncFlow(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }), deps),
    /Remote snapshot failed on media-box[\s\S]*snapshot permission denied/,
  );
});

test('runHostSync push only snapshots locally and merges remotely', async () => {
  const calls: string[] = [];
  await runSyncFlow(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box', syncDirection: 'push' }),
    makeHostDeps(calls),
  );

  assert.ok(calls.some((call) => call.startsWith('snapshot:')));
  assert.ok(calls.some((call) => call.includes(' sync --merge ')));
  assert.ok(calls.some((call) => call.startsWith('scp:') && call.includes('->media-box:')));
  assert.ok(!calls.some((call) => call.includes(' sync --snapshot ')));
  assert.ok(!calls.includes('local-merge'));
});

test('runHostSync pull only snapshots remotely and merges locally', async () => {
  const calls: string[] = [];
  await runSyncFlow(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box', syncDirection: 'pull' }),
    makeHostDeps(calls),
  );

  assert.ok(calls.some((call) => call.includes(' sync --snapshot ')));
  assert.ok(calls.some((call) => call.startsWith('scp:media-box:')));
  assert.ok(calls.includes('local-merge'));
  assert.ok(!calls.some((call) => call.startsWith('snapshot:')));
  assert.ok(!calls.some((call) => call.includes(' sync --merge ')));
});

test('runSyncFlow --json emits NDJSON progress events and a final result', async () => {
  const lines: string[] = [];
  const remoteSummary = {
    ...createEmptyMergeSummary(),
    sessionsMerged: 2,
    videosAdded: 1,
  };
  const deps = makeHostDeps([], {
    consoleLog: (line) => {
      lines.push(line);
    },
    runSsh: (_host, command) => {
      if (command.includes(' sync --make-temp')) return ok('/tmp/subminer-sync-remote\n');
      if (command.includes(' sync --merge ')) {
        assert.match(command, / --json(?: |$)/);
        return ok(
          `${JSON.stringify({ type: 'merge-summary', target: 'local', summary: remoteSummary })}\n` +
            `${JSON.stringify({ type: 'result', ok: true, error: null })}\n`,
        );
      }
      return ok();
    },
  });

  await runSyncFlow(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box', syncJson: true }),
    deps,
  );

  const events = lines.map((line) => JSON.parse(line));
  assert.ok(events.some((event) => event.type === 'stage' && event.stage === 'snapshot-local'));
  assert.ok(events.some((event) => event.type === 'merge-summary' && event.target === 'local'));
  assert.deepEqual(
    events.find((event) => event.type === 'merge-summary' && event.target === 'remote'),
    { type: 'merge-summary', target: 'remote', summary: remoteSummary },
  );
  assert.deepEqual(events[events.length - 1], { type: 'result', ok: true, error: null });
});

test('runSyncFlow --json emits an error result when the sync fails', async () => {
  const lines: string[] = [];
  const deps = makeHostDeps([], {
    consoleLog: (line) => {
      lines.push(line);
    },
    runSsh: (_host, command) => {
      if (command.includes(' sync --make-temp')) return ok('/tmp/subminer-sync-remote\n');
      if (command.includes(' sync --merge ')) return { status: 9, stdout: '', stderr: 'boom' };
      return ok();
    },
  });

  await assert.rejects(() =>
    runSyncFlow(
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

test('runHostSync records host sync results for saved-host bookkeeping', async () => {
  const recorded: Array<{ host: string; status: string; detail: string | null }> = [];
  const record: SyncFlowDeps['recordHostSyncResult'] = (host, status, detail) => {
    recorded.push({ host, status, detail });
  };

  await runSyncFlow(
    makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }),
    makeHostDeps([], { recordHostSyncResult: record }),
  );
  assert.deepEqual(recorded[0], {
    host: 'media-box',
    status: 'success',
    detail: '0 sessions merged; pushed local stats',
  });

  await assert.rejects(() =>
    runSyncFlow(
      makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'media-box' }),
      makeHostDeps([], {
        recordHostSyncResult: record,
        runSsh: (_host, command) => {
          if (command.includes(' sync --make-temp')) return ok('/tmp/subminer-sync-remote\n');
          if (command.includes(' sync --merge ')) return { status: 9, stdout: '', stderr: 'boom' };
          return ok();
        },
      }),
    ),
  );
  assert.equal(recorded.length, 2);
  assert.equal(recorded[1]!.status, 'error');
});

test('runCheckMode --json reports ssh and remote SubMiner status', async () => {
  const lines: string[] = [];
  const sshOptions: unknown[] = [];
  const deps = makeDeps({
    consoleLog: (line) => {
      lines.push(line);
    },
    resolveRemoteSubminerCommand: (host, _preferred, _flavor, runRemote) => {
      runRemote!(host, 'subminer --help');
      return 'subminer';
    },
    runSsh: (_host, command, options) => {
      sshOptions.push(options);
      if (command.includes('--version')) return ok('SubMiner 0.18.0\n');
      return ok('subminer-check-ok\n');
    },
  });

  await runSyncFlow(
    makeContext({
      syncDbPath: '/tmp/local.sqlite',
      syncHost: 'media-box',
      syncCheck: true,
      syncJson: true,
    }),
    deps,
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
  assert.equal(sshOptions.length, 3);
  assert.ok(
    sshOptions.every(
      (options) =>
        JSON.stringify(options) ===
        JSON.stringify({ batchMode: true, connectTimeoutSeconds: 10, timeoutMs: 15_000 }),
    ),
  );

  const failLines: string[] = [];
  await assert.rejects(() =>
    runSyncFlow(
      makeContext({
        syncDbPath: '/tmp/local.sqlite',
        syncHost: 'media-box',
        syncCheck: true,
        syncJson: true,
      }),
      makeDeps({
        consoleLog: (line) => {
          failLines.push(line);
        },
        resolveRemoteSubminerCommand: () => {
          throw new Error('Could not find a runnable "subminer" on media-box.');
        },
        runSsh: () => ok('subminer-check-ok\n'),
      }),
    ),
  );
  const failEvents = failLines.map((line) => JSON.parse(line));
  const failedCheck = failEvents.find((event) => event.type === 'check-result');
  assert.ok(failedCheck);
  assert.equal(failedCheck.ok, false);
  assert.match(failedCheck.error, /Could not find a runnable/);
});

test('runHostSync speaks Windows shells: app command, double quotes, temp protocol', async () => {
  const sshCommands: string[] = [];
  const scpCalls: string[] = [];
  const winTemp = 'C:\\Users\\First Last\\AppData\\Local\\Temp\\subminer-sync-remote';
  const appCmd = '"%LOCALAPPDATA%\\Programs\\SubMiner\\SubMiner.exe" --sync-cli';
  const deps = makeHostDeps([], {
    detectRemoteShellFlavor: () => 'windows-cmd',
    resolveRemoteSubminerCommand: () => appCmd,
    createDbSnapshot: (_dbPath, outPath) => {
      fs.writeFileSync(outPath, 'snapshot');
    },
    runSsh: (_host, command) => {
      sshCommands.push(command);
      if (command.includes(' sync --make-temp')) return ok(`${winTemp}\r\n`);
      return ok();
    },
    runScp: (from, to) => {
      scpCalls.push(`${from}->${to}`);
      if (!to.includes(':')) fs.writeFileSync(to, 'pulled');
    },
  });

  await runSyncFlow(makeContext({ syncDbPath: '/tmp/local.sqlite', syncHost: 'win-box' }), deps);

  const expectedDir = 'C:/Users/First Last/AppData/Local/Temp/subminer-sync-remote';
  assert.ok(sshCommands.includes(`${appCmd} sync --snapshot "${expectedDir}/snapshot.sqlite"`));
  assert.ok(sshCommands.includes(`${appCmd} sync --merge "${expectedDir}/incoming.sqlite"`));
  assert.ok(sshCommands.includes(`${appCmd} sync --remove-temp "${expectedDir}"`));
  assert.ok(scpCalls.some((call) => call.startsWith(`win-box:${expectedDir}/snapshot.sqlite->`)));
  assert.ok(scpCalls.some((call) => call.endsWith(`->win-box:${expectedDir}/incoming.sqlite`)));
  // No POSIX shell-isms reach a Windows remote.
  assert.ok(
    !sshCommands.some((command) => command.startsWith('mktemp') || command.startsWith('rm ')),
  );
  assert.ok(!sshCommands.some((command) => command.includes('PATH="$HOME')));
});

test('runSyncFlow --make-temp prints a temp dir and --remove-temp only removes sync temp dirs', async () => {
  const printed: string[] = [];
  const removed: string[] = [];
  const madeDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-sync-'));
  try {
    await runSyncFlow(
      makeContext({ syncMakeTemp: true }),
      makeDeps({
        mkdtempSync: () => madeDir,
        consoleLog: (line) => {
          printed.push(line);
        },
      }),
    );
    assert.deepEqual(printed, [madeDir]);

    const removeDeps = makeDeps({
      rmSync: (target) => {
        removed.push(target);
      },
    });
    await runSyncFlow(makeContext({ syncRemoveTempPath: madeDir }), removeDeps);
    assert.deepEqual(removed, [madeDir]);

    await assert.rejects(
      () => runSyncFlow(makeContext({ syncRemoveTempPath: '/etc' }), removeDeps),
      /Refusing to remove a directory outside the sync temp area/,
    );
    assert.equal(removed.length, 1);
  } finally {
    fs.rmSync(madeDir, { recursive: true, force: true });
  }
});
