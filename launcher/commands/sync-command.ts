import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fail, log } from '../log.js';
import { resolveImmersionDbPath } from '../history-db.js';
import {
  createDbSnapshot,
  findLiveStatsDaemonPid,
  formatMergeSummary,
  mergeSnapshotIntoDb,
} from '../sync/sync-db.js';
import {
  assertSafeSshHost,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
  shellQuote,
} from '../sync/ssh.js';
import { resolvePathMaybe } from '../util.js';
import { canConnectUnixSocket } from '../mpv.js';
import type { LauncherCommandContext } from './context.js';
import type { RemoteRunResult } from '../sync/ssh.js';

export interface SyncCommandDeps {
  createDbSnapshot: typeof createDbSnapshot;
  mergeSnapshotIntoDb: typeof mergeSnapshotIntoDb;
  formatMergeSummary: typeof formatMergeSummary;
  findLiveStatsDaemonPid: typeof findLiveStatsDaemonPid;
  assertSafeSshHost: typeof assertSafeSshHost;
  resolveRemoteSubminerCommand: typeof resolveRemoteSubminerCommand;
  runScp: typeof runScp;
  runSsh: typeof runSsh;
  fail: typeof fail;
  log: typeof log;
  canConnectUnixSocket: typeof canConnectUnixSocket;
  realpathSync: typeof fs.realpathSync;
  mkdtempSync: typeof fs.mkdtempSync;
  rmSync: typeof fs.rmSync;
  consoleLog: typeof console.log;
  writeStdout: typeof process.stdout.write;
  ensureTrackerQuiescent: (context: LauncherCommandContext, dbPath: string) => Promise<void>;
}

function resolveDbPath(context: LauncherCommandContext): string {
  const override = context.args.syncDbPath.trim();
  return override ? resolvePathMaybe(override) : resolveImmersionDbPath();
}

function isTrackerDb(dbPath: string, deps: SyncCommandDeps): boolean {
  const trackerDbPath = resolveImmersionDbPath();
  try {
    return deps.realpathSync(dbPath) === deps.realpathSync(trackerDbPath);
  } catch {
    return dbPath === trackerDbPath;
  }
}

export async function ensureTrackerQuiescent(
  context: LauncherCommandContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<void> {
  const deps = resolveSyncCommandDeps(inputDeps);
  if (context.args.syncForce) return;
  // A running SubMiner only holds the tracker's own database; --db pointed
  // elsewhere needs no guard.
  if (!isTrackerDb(dbPath, deps)) return;
  const daemonPid = deps.findLiveStatsDaemonPid(dbPath);
  if (daemonPid !== null) {
    deps.fail(
      `The SubMiner stats server is running (pid ${daemonPid}). Stop it with "subminer stats -s" (or close SubMiner) before syncing, or pass --force.`,
    );
  }
  if (context.mpvSocketPath && (await deps.canConnectUnixSocket(context.mpvSocketPath))) {
    deps.fail(
      `An mpv/SubMiner session appears to be running (socket ${context.mpvSocketPath}). Close it before syncing, or pass --force.`,
    );
  }
}

const defaultSyncCommandDeps: SyncCommandDeps = {
  createDbSnapshot,
  mergeSnapshotIntoDb,
  formatMergeSummary,
  findLiveStatsDaemonPid,
  assertSafeSshHost,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
  fail,
  log,
  canConnectUnixSocket,
  realpathSync: fs.realpathSync,
  mkdtempSync: fs.mkdtempSync,
  rmSync: fs.rmSync,
  consoleLog: console.log,
  writeStdout: process.stdout.write.bind(process.stdout),
  ensureTrackerQuiescent: async (context, dbPath) => ensureTrackerQuiescent(context, dbPath),
};

function resolveSyncCommandDeps(inputDeps: Partial<SyncCommandDeps> = {}): SyncCommandDeps {
  return { ...defaultSyncCommandDeps, ...inputDeps };
}

export function runSnapshotMode(
  context: LauncherCommandContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): void {
  const deps = resolveSyncCommandDeps(inputDeps);
  const outPath = resolvePathMaybe(context.args.syncSnapshotPath);
  deps.createDbSnapshot(dbPath, outPath);
  deps.consoleLog(outPath);
}

export async function runMergeMode(
  context: LauncherCommandContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<void> {
  const deps = resolveSyncCommandDeps(inputDeps);
  await deps.ensureTrackerQuiescent(context, dbPath);
  const snapshotPath = resolvePathMaybe(context.args.syncMergePath);
  const summary = deps.mergeSnapshotIntoDb(dbPath, snapshotPath);
  deps.consoleLog(deps.formatMergeSummary(summary));
}

function cleanupRemote(host: string, remoteTmpDir: string, deps: SyncCommandDeps): void {
  if (!remoteTmpDir.startsWith('/tmp/')) return;
  deps.runSsh(host, `rm -rf ${shellQuote(remoteTmpDir)}`);
}

function formatRemoteRunError(message: string, run: RemoteRunResult): string {
  const stderr = run.stderr.trim();
  return stderr ? `${message}\n${stderr}` : message;
}

export async function runHostSync(
  context: LauncherCommandContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<void> {
  const deps = resolveSyncCommandDeps(inputDeps);
  const { args } = context;
  const host = args.syncHost;
  const direction = args.syncDirection ?? 'both';
  const shouldPull = direction !== 'push';
  const shouldPush = direction !== 'pull';
  deps.assertSafeSshHost(host);

  await deps.ensureTrackerQuiescent(context, dbPath);

  const remoteCmd = deps.resolveRemoteSubminerCommand(host, args.syncRemoteCmd || null);
  deps.log('debug', args.logLevel, `Remote subminer command: ${remoteCmd}`);

  const localTmpDir = deps.mkdtempSync(path.join(os.tmpdir(), 'subminer-sync-'));
  let remoteTmpDir = '';
  try {
    // Signal failures by throwing (not fail(), which exits synchronously and
    // would skip the finally cleanup, leaking temp dirs holding snapshot data).
    // main().catch() reports the message the same way fail() would.
    const mktemp = deps.runSsh(host, 'mktemp -d /tmp/subminer-sync.XXXXXX');
    remoteTmpDir = mktemp.stdout.trim();
    if (mktemp.status !== 0 || !remoteTmpDir.startsWith('/tmp/')) {
      throw new Error(`Could not create a temporary directory on ${host}.`);
    }

    const forceFlag = args.syncForce ? ' --force' : '';

    const localSnapshot = path.join(localTmpDir, 'local.sqlite');
    if (shouldPush) {
      deps.consoleLog(`Snapshotting local database (${dbPath})...`);
      deps.createDbSnapshot(dbPath, localSnapshot);
    }

    const remoteSnapshot = `${remoteTmpDir}/snapshot.sqlite`;
    if (shouldPull) {
      deps.consoleLog(`Snapshotting ${host}...`);
      const snapshotRun = deps.runSsh(
        host,
        `${remoteCmd} sync --snapshot ${shellQuote(remoteSnapshot)}${forceFlag}`,
      );
      if (snapshotRun.status !== 0) {
        throw new Error(formatRemoteRunError(`Remote snapshot failed on ${host}.`, snapshotRun));
      }
    }

    const pulledSnapshot = path.join(localTmpDir, 'remote.sqlite');
    if (shouldPull) deps.runScp(`${host}:${remoteSnapshot}`, pulledSnapshot);
    const incomingSnapshot = `${remoteTmpDir}/incoming.sqlite`;
    if (shouldPush) deps.runScp(localSnapshot, `${host}:${incomingSnapshot}`);

    if (shouldPull) {
      deps.consoleLog(`\nMerging ${host} -> local:`);
      await deps.ensureTrackerQuiescent(context, dbPath);
      const summary = deps.mergeSnapshotIntoDb(dbPath, pulledSnapshot);
      deps.consoleLog(deps.formatMergeSummary(summary));
    }

    if (shouldPush) {
      deps.consoleLog(`\nMerging local -> ${host}:`);
      await deps.ensureTrackerQuiescent(context, dbPath);
      const mergeRun = deps.runSsh(
        host,
        `${remoteCmd} sync --merge ${shellQuote(incomingSnapshot)}${forceFlag}`,
      );
      deps.writeStdout(mergeRun.stdout);
      if (mergeRun.status !== 0) {
        const retryCommand =
          direction === 'push' ? `subminer sync ${host} --push` : `subminer sync ${host}`;
        const localUpdate = shouldPull ? ' The local database was updated;' : '';
        throw new Error(
          formatRemoteRunError(
            `Remote merge failed on ${host}.${localUpdate} re-run "${retryCommand}" once the remote issue is fixed.`,
            mergeRun,
          ),
        );
      }
    }

    deps.consoleLog('\nSync complete.');
  } finally {
    deps.rmSync(localTmpDir, { recursive: true, force: true });
    if (remoteTmpDir) {
      try {
        cleanupRemote(host, remoteTmpDir, deps);
      } catch {
        // best effort
      }
    }
  }
}

export async function runSyncCommand(
  context: LauncherCommandContext,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<boolean> {
  const deps = resolveSyncCommandDeps(inputDeps);
  const { args } = context;
  if (!args.sync) return false;

  const dbPath = resolveDbPath(context);
  if (args.syncSnapshotPath) {
    runSnapshotMode(context, dbPath, deps);
  } else if (args.syncMergePath) {
    await runMergeMode(context, dbPath, deps);
  } else if (args.syncHost) {
    await runHostSync(context, dbPath, deps);
  } else {
    deps.fail('sync requires a host, --snapshot <file>, or --merge <file>.');
  }
  return true;
}
