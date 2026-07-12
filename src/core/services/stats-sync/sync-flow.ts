import os from 'node:os';
import path from 'node:path';
import { shellQuote } from './ssh';
import type { RemoteRunResult } from './ssh';
import type { SyncMergeSummary, SyncProgressEvent } from '../../../shared/sync/sync-events';
import type { SyncResultStatus } from '../../../shared/sync/sync-hosts-store';

export interface SyncFlowArgs {
  sync: boolean;
  syncHost: string;
  syncSnapshotPath: string;
  syncMergePath: string;
  syncDirection: 'both' | 'push' | 'pull' | null;
  syncRemoteCmd: string;
  syncDbPath: string;
  syncForce: boolean;
  syncJson: boolean;
  syncCheck: boolean;
  logLevel: string;
}

export interface SyncFlowContext {
  args: SyncFlowArgs;
  mpvSocketPath: string;
}

/**
 * Everything the sync flow touches outside plain computation. The launcher
 * binds these to bun:sqlite-backed engine functions and its own logger; the
 * app's --sync-cli mode binds libsql and console output. Tests stub freely.
 */
export interface SyncFlowDeps {
  createDbSnapshot: (dbPath: string, outPath: string) => void;
  mergeSnapshotIntoDb: (localDbPath: string, snapshotPath: string) => SyncMergeSummary;
  formatMergeSummary: (summary: SyncMergeSummary) => string;
  findLiveStatsDaemonPid: (dbPath: string) => number | null;
  assertSafeSshHost: (host: string) => void;
  resolveRemoteSubminerCommand: (host: string, preferred: string | null) => string;
  runScp: (from: string, to: string) => void;
  runSsh: (host: string, remoteCommand: string) => RemoteRunResult;
  fail: (message: string) => never;
  log: (level: string, configuredLevel: string, message: string) => void;
  canConnectUnixSocket: (socketPath: string) => Promise<boolean>;
  realpathSync: (candidate: string) => string;
  mkdtempSync: (prefix: string) => string;
  rmSync: (target: string, options: { recursive: boolean; force: boolean }) => void;
  consoleLog: (message: string) => void;
  writeStdout: (text: string) => boolean;
  ensureTrackerQuiescent: (context: SyncFlowContext, dbPath: string) => Promise<void>;
  emitEvent: (event: SyncProgressEvent) => void;
  recordHostSyncResult: (host: string, status: SyncResultStatus, detail: string | null) => void;
  resolveDefaultDbPath: () => string;
  resolvePath: (value: string) => string;
}

export function resolveSyncDbPath(context: SyncFlowContext, deps: SyncFlowDeps): string {
  const override = context.args.syncDbPath.trim();
  return override ? deps.resolvePath(override) : deps.resolveDefaultDbPath();
}

function isTrackerDb(dbPath: string, deps: SyncFlowDeps): boolean {
  const trackerDbPath = deps.resolveDefaultDbPath();
  try {
    return deps.realpathSync(dbPath) === deps.realpathSync(trackerDbPath);
  } catch {
    return dbPath === trackerDbPath;
  }
}

export async function ensureTrackerQuiescentFlow(
  context: SyncFlowContext,
  dbPath: string,
  deps: SyncFlowDeps,
): Promise<void> {
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

// In --json mode every line on stdout is an NDJSON event: human console output
// is silenced and events are written through the original console logger.
export function withJsonEvents(deps: SyncFlowDeps): SyncFlowDeps {
  const writeLine = deps.consoleLog;
  return {
    ...deps,
    consoleLog: () => {},
    writeStdout: () => true,
    emitEvent: (event) => writeLine(JSON.stringify(event)),
  };
}

export function runSnapshotMode(
  context: SyncFlowContext,
  dbPath: string,
  deps: SyncFlowDeps,
): void {
  const outPath = deps.resolvePath(context.args.syncSnapshotPath);
  deps.emitEvent({
    type: 'stage',
    stage: 'snapshot-local',
    message: `Snapshotting local database (${dbPath})`,
  });
  deps.createDbSnapshot(dbPath, outPath);
  deps.emitEvent({ type: 'snapshot-created', path: outPath });
  deps.consoleLog(outPath);
}

export async function runMergeMode(
  context: SyncFlowContext,
  dbPath: string,
  deps: SyncFlowDeps,
): Promise<void> {
  await deps.ensureTrackerQuiescent(context, dbPath);
  const snapshotPath = deps.resolvePath(context.args.syncMergePath);
  deps.emitEvent({
    type: 'stage',
    stage: 'merge-local',
    message: `Merging ${snapshotPath} into the local database`,
  });
  const summary = deps.mergeSnapshotIntoDb(dbPath, snapshotPath);
  deps.emitEvent({ type: 'merge-summary', target: 'local', summary });
  deps.consoleLog(deps.formatMergeSummary(summary));
}

function formatHostSyncDetail(
  direction: 'both' | 'push' | 'pull',
  pulledSummary: SyncMergeSummary | null,
): string {
  if (!pulledSummary) return direction === 'push' ? 'Pushed local stats' : 'Sync complete';
  const merged = `${pulledSummary.sessionsMerged} session${pulledSummary.sessionsMerged === 1 ? '' : 's'} merged`;
  return direction === 'pull' ? merged : `${merged}; pushed local stats`;
}

export async function runCheckMode(context: SyncFlowContext, deps: SyncFlowDeps): Promise<void> {
  const { args } = context;
  const host = args.syncHost;
  deps.assertSafeSshHost(host);

  deps.consoleLog(`Checking SSH connection to ${host}...`);
  let remoteCommand: string | null = null;
  let remoteVersion: string | null = null;
  let error: string | null = null;

  const probe = deps.runSsh(host, 'echo subminer-check-ok');
  const sshOk = probe.status === 0 && probe.stdout.includes('subminer-check-ok');
  if (!sshOk) {
    error = formatRemoteRunError(`Could not reach ${host} over SSH.`, probe);
  } else {
    deps.consoleLog('SSH connection: ok');
    try {
      remoteCommand = deps.resolveRemoteSubminerCommand(host, args.syncRemoteCmd || null);
      const version = deps.runSsh(host, `${remoteCommand} --version`);
      remoteVersion = version.status === 0 ? version.stdout.trim() || null : null;
      deps.consoleLog(
        `Remote subminer: ${remoteCommand}${remoteVersion ? ` (${remoteVersion})` : ''}`,
      );
    } catch (resolveError) {
      error = resolveError instanceof Error ? resolveError.message : String(resolveError);
    }
  }

  const ok = sshOk && remoteCommand !== null;
  deps.emitEvent({
    type: 'check-result',
    host,
    sshOk,
    remoteCommand,
    remoteVersion,
    ok,
    error,
  });
  if (!ok) {
    throw new Error(error ?? `Connection check failed for ${host}.`);
  }
  deps.consoleLog('Check passed.');
}

function cleanupRemote(host: string, remoteTmpDir: string, deps: SyncFlowDeps): void {
  if (!remoteTmpDir.startsWith('/tmp/')) return;
  deps.runSsh(host, `rm -rf ${shellQuote(remoteTmpDir)}`);
}

function formatRemoteRunError(message: string, run: RemoteRunResult): string {
  const stderr = run.stderr.trim();
  return stderr ? `${message}\n${stderr}` : message;
}

export async function runHostSync(
  context: SyncFlowContext,
  dbPath: string,
  deps: SyncFlowDeps,
): Promise<void> {
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
  let pulledSummary: SyncMergeSummary | null = null;
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
      deps.emitEvent({
        type: 'stage',
        stage: 'snapshot-local',
        message: `Snapshotting local database (${dbPath})`,
      });
      deps.createDbSnapshot(dbPath, localSnapshot);
    }

    const remoteSnapshot = `${remoteTmpDir}/snapshot.sqlite`;
    if (shouldPull) {
      deps.consoleLog(`Snapshotting ${host}...`);
      deps.emitEvent({ type: 'stage', stage: 'snapshot-remote', message: `Snapshotting ${host}` });
      const snapshotRun = deps.runSsh(
        host,
        `${remoteCmd} sync --snapshot ${shellQuote(remoteSnapshot)}${forceFlag}`,
      );
      if (snapshotRun.status !== 0) {
        throw new Error(formatRemoteRunError(`Remote snapshot failed on ${host}.`, snapshotRun));
      }
    }

    const pulledSnapshot = path.join(localTmpDir, 'remote.sqlite');
    if (shouldPull) {
      deps.emitEvent({
        type: 'stage',
        stage: 'download',
        message: `Copying snapshot from ${host}`,
      });
      deps.runScp(`${host}:${remoteSnapshot}`, pulledSnapshot);
    }
    const incomingSnapshot = `${remoteTmpDir}/incoming.sqlite`;
    if (shouldPush) {
      deps.emitEvent({ type: 'stage', stage: 'upload', message: `Copying snapshot to ${host}` });
      deps.runScp(localSnapshot, `${host}:${incomingSnapshot}`);
    }

    if (shouldPull) {
      deps.consoleLog(`\nMerging ${host} -> local:`);
      deps.emitEvent({
        type: 'stage',
        stage: 'merge-local',
        message: `Merging ${host} into the local database`,
      });
      await deps.ensureTrackerQuiescent(context, dbPath);
      const summary = deps.mergeSnapshotIntoDb(dbPath, pulledSnapshot);
      pulledSummary = summary;
      deps.emitEvent({ type: 'merge-summary', target: 'local', summary });
      deps.consoleLog(deps.formatMergeSummary(summary));
    }

    if (shouldPush) {
      deps.consoleLog(`\nMerging local -> ${host}:`);
      deps.emitEvent({
        type: 'stage',
        stage: 'merge-remote',
        message: `Merging the local database into ${host}`,
      });
      await deps.ensureTrackerQuiescent(context, dbPath);
      const mergeRun = deps.runSsh(
        host,
        `${remoteCmd} sync --merge ${shellQuote(incomingSnapshot)}${forceFlag}`,
      );
      deps.writeStdout(mergeRun.stdout);
      if (mergeRun.stdout.trim()) {
        deps.emitEvent({ type: 'remote-output', text: mergeRun.stdout });
      }
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
    deps.recordHostSyncResult(host, 'success', formatHostSyncDetail(direction, pulledSummary));
  } catch (error) {
    try {
      deps.recordHostSyncResult(
        host,
        'error',
        error instanceof Error ? error.message : String(error),
      );
    } catch {
      // best effort
    }
    throw error;
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

export async function runSyncFlow(context: SyncFlowContext, inputDeps: SyncFlowDeps): Promise<boolean> {
  let deps = inputDeps;
  const { args } = context;
  if (!args.sync) return false;
  if (args.syncJson) deps = withJsonEvents(deps);

  const dbPath = resolveSyncDbPath(context, deps);
  try {
    if (args.syncCheck) {
      await runCheckMode(context, deps);
    } else if (args.syncSnapshotPath) {
      runSnapshotMode(context, dbPath, deps);
    } else if (args.syncMergePath) {
      await runMergeMode(context, dbPath, deps);
    } else if (args.syncHost) {
      await runHostSync(context, dbPath, deps);
    } else {
      deps.fail('sync requires a host, --snapshot <file>, or --merge <file>.');
    }
  } catch (error) {
    if (args.syncJson) {
      deps.emitEvent({
        type: 'result',
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    }
    throw error;
  }
  if (args.syncJson) deps.emitEvent({ type: 'result', ok: true, error: null });
  return true;
}
