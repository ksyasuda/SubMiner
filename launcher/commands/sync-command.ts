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
import type { LauncherCommandContext } from './context.js';

function resolveDbPath(context: LauncherCommandContext): string {
  const override = context.args.syncDbPath.trim();
  return override ? resolvePathMaybe(override) : resolveImmersionDbPath();
}

function isTrackerDb(dbPath: string): boolean {
  const trackerDbPath = resolveImmersionDbPath();
  try {
    return fs.realpathSync(dbPath) === fs.realpathSync(trackerDbPath);
  } catch {
    return dbPath === trackerDbPath;
  }
}

function ensureTrackerQuiescent(context: LauncherCommandContext, dbPath: string): void {
  if (context.args.syncForce) return;
  // A running SubMiner only holds the tracker's own database; --db pointed
  // elsewhere needs no guard.
  if (!isTrackerDb(dbPath)) return;
  const daemonPid = findLiveStatsDaemonPid(dbPath);
  if (daemonPid !== null) {
    fail(
      `The SubMiner stats server is running (pid ${daemonPid}). Stop it with "subminer stats -s" (or close SubMiner) before syncing, or pass --force.`,
    );
  }
  if (context.mpvSocketPath && fs.existsSync(context.mpvSocketPath)) {
    fail(
      `An mpv/SubMiner session appears to be running (socket ${context.mpvSocketPath}). Close it before syncing, or pass --force.`,
    );
  }
}

function runSnapshotMode(context: LauncherCommandContext, dbPath: string): void {
  const outPath = resolvePathMaybe(context.args.syncSnapshotPath);
  createDbSnapshot(dbPath, outPath);
  console.log(outPath);
}

function runMergeMode(context: LauncherCommandContext, dbPath: string): void {
  ensureTrackerQuiescent(context, dbPath);
  const snapshotPath = resolvePathMaybe(context.args.syncMergePath);
  const summary = mergeSnapshotIntoDb(dbPath, snapshotPath);
  console.log(formatMergeSummary(summary));
}

function cleanupRemote(host: string, remoteTmpDir: string): void {
  if (!remoteTmpDir.startsWith('/tmp/')) return;
  runSsh(host, `rm -rf ${shellQuote(remoteTmpDir)}`);
}

function runHostSync(context: LauncherCommandContext, dbPath: string): void {
  const { args } = context;
  const host = args.syncHost;
  assertSafeSshHost(host);

  ensureTrackerQuiescent(context, dbPath);

  const remoteCmd = resolveRemoteSubminerCommand(host, args.syncRemoteCmd || null);
  log('debug', args.logLevel, `Remote subminer command: ${remoteCmd}`);

  const localTmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-sync-'));
  let remoteTmpDir = '';
  try {
    // Signal failures by throwing (not fail(), which exits synchronously and
    // would skip the finally cleanup, leaking temp dirs holding snapshot data).
    // main().catch() reports the message the same way fail() would.
    const mktemp = runSsh(host, 'mktemp -d /tmp/subminer-sync.XXXXXX');
    remoteTmpDir = mktemp.stdout.trim();
    if (mktemp.status !== 0 || !remoteTmpDir.startsWith('/tmp/')) {
      throw new Error(`Could not create a temporary directory on ${host}.`);
    }

    const forceFlag = args.syncForce ? ' --force' : '';

    console.log(`Snapshotting local database (${dbPath})...`);
    const localSnapshot = path.join(localTmpDir, 'local.sqlite');
    createDbSnapshot(dbPath, localSnapshot);

    console.log(`Snapshotting ${host}...`);
    const remoteSnapshot = `${remoteTmpDir}/snapshot.sqlite`;
    const snapshotRun = runSsh(
      host,
      `${remoteCmd} sync --snapshot ${shellQuote(remoteSnapshot)}${forceFlag}`,
    );
    if (snapshotRun.status !== 0) {
      throw new Error(`Remote snapshot failed on ${host}.`);
    }

    const pulledSnapshot = path.join(localTmpDir, 'remote.sqlite');
    runScp(`${host}:${remoteSnapshot}`, pulledSnapshot);
    const incomingSnapshot = `${remoteTmpDir}/incoming.sqlite`;
    runScp(localSnapshot, `${host}:${incomingSnapshot}`);

    console.log(`\nMerging ${host} -> local:`);
    const summary = mergeSnapshotIntoDb(dbPath, pulledSnapshot);
    console.log(formatMergeSummary(summary));

    console.log(`\nMerging local -> ${host}:`);
    const mergeRun = runSsh(
      host,
      `${remoteCmd} sync --merge ${shellQuote(incomingSnapshot)}${forceFlag}`,
    );
    process.stdout.write(mergeRun.stdout);
    if (mergeRun.status !== 0) {
      throw new Error(
        `Remote merge failed on ${host}. The local database was updated; re-run "subminer sync ${host}" once the remote issue is fixed.`,
      );
    }

    console.log('\nSync complete.');
  } finally {
    fs.rmSync(localTmpDir, { recursive: true, force: true });
    if (remoteTmpDir) {
      try {
        cleanupRemote(host, remoteTmpDir);
      } catch {
        // best effort
      }
    }
  }
}

export function runSyncCommand(context: LauncherCommandContext): boolean {
  const { args } = context;
  if (!args.sync) return false;

  const dbPath = resolveDbPath(context);
  if (args.syncSnapshotPath) {
    runSnapshotMode(context, dbPath);
  } else if (args.syncMergePath) {
    runMergeMode(context, dbPath);
  } else if (args.syncHost) {
    runHostSync(context, dbPath);
  } else {
    fail('sync requires a host, --snapshot <file>, or --merge <file>.');
  }
  return true;
}
