import os from 'node:os';
import path from 'node:path';
import { formatMergeSummary } from './merge';
import { quoteForRemoteShell } from './ssh';
import type { RemoteRunResult, RemoteShellFlavor, RunSshOptions } from './ssh';
import {
  parseSyncProgressLine,
  type SyncMergeSummary,
  type SyncProgressEvent,
} from '../../../shared/sync/sync-events';
import type { SyncResultStatus } from '../../../shared/sync/sync-hosts-store';

export interface SyncFlowArgs {
  syncHost: string;
  syncSnapshotPath: string;
  syncMergePath: string;
  syncDirection: 'both' | 'push' | 'pull' | null;
  syncRemoteCmd: string;
  syncDbPath: string;
  syncForce: boolean;
  syncJson: boolean;
  syncCheck: boolean;
  syncMakeTemp: boolean;
  syncRemoveTempPath: string;
  logLevel: string;
}

export interface SyncFlowContext {
  args: SyncFlowArgs;
  mpvSocketPath: string;
}

/**
 * Process/IO seams the sync flow needs stubbed in tests: SSH/scp, the DB
 * snapshot/merge engine, filesystem, and progress/bookkeeping output. The
 * app's --sync-cli mode (src/main/sync-cli.ts) provides the only production
 * binding; pure helpers are imported directly.
 */
export interface SyncFlowDeps {
  createDbSnapshot: (dbPath: string, outPath: string) => void;
  mergeSnapshotIntoDb: (localDbPath: string, snapshotPath: string) => SyncMergeSummary;
  findLiveStatsDaemonPid: (dbPath: string) => number | null;
  assertSafeSshHost: (host: string) => void;
  detectRemoteShellFlavor: (
    host: string,
    runRemote: (host: string, remoteCommand: string) => RemoteRunResult,
  ) => RemoteShellFlavor;
  resolveRemoteSubminerCommand: (
    host: string,
    preferred: string | null,
    flavor: RemoteShellFlavor,
    runRemote?: (host: string, remoteCommand: string) => RemoteRunResult,
  ) => string;
  runScp: (from: string, to: string) => void;
  runSsh: (host: string, remoteCommand: string, options?: RunSshOptions) => RemoteRunResult;
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
}

/** Expand a leading `~` (as ssh users write paths) and make the path absolute. */
function resolveCliPath(input: string): string {
  return input.startsWith('~') ? path.join(os.homedir(), input.slice(1)) : path.resolve(input);
}

/**
 * Prefix shared by every sync temp dir, local or remote. Remote temp dirs are
 * created and removed by the remote SubMiner itself (sync --make-temp /
 * --remove-temp) so the flow never depends on mktemp/rm existing in the
 * remote shell. This is what makes Windows remotes work.
 */
const SYNC_TEMP_PREFIX = 'subminer-sync-';

function makeSyncTempDir(mkdtempSync: SyncFlowDeps['mkdtempSync']): string {
  return mkdtempSync(path.join(os.tmpdir(), SYNC_TEMP_PREFIX));
}

/** Only dirs directly under os.tmpdir() with the sync prefix may be removed. */
function assertRemovableSyncTempDir(target: string): string {
  const resolved = path.resolve(target.trim());
  const normalizeCase = (value: string) =>
    process.platform === 'win32' ? value.toLowerCase() : value;
  const insideTmp =
    normalizeCase(path.dirname(resolved)) === normalizeCase(path.resolve(os.tmpdir()));
  if (!insideTmp || !path.basename(resolved).startsWith(SYNC_TEMP_PREFIX)) {
    throw new Error(`Refusing to remove a directory outside the sync temp area: ${target}`);
  }
  return resolved;
}

function runMakeTempMode(deps: SyncFlowDeps): void {
  deps.consoleLog(makeSyncTempDir(deps.mkdtempSync));
}

function runRemoveTempMode(context: SyncFlowContext, deps: SyncFlowDeps): void {
  const target = assertRemovableSyncTempDir(context.args.syncRemoveTempPath);
  deps.rmSync(target, { recursive: true, force: true });
}

function resolveSyncDbPath(context: SyncFlowContext, deps: SyncFlowDeps): string {
  const override = context.args.syncDbPath.trim();
  return override ? resolveCliPath(override) : deps.resolveDefaultDbPath();
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
    throw new Error(
      `The SubMiner stats server is running (pid ${daemonPid}). Stop it with "subminer stats -s" (or close SubMiner) before syncing, or pass --force.`,
    );
  }
  if (context.mpvSocketPath && (await deps.canConnectUnixSocket(context.mpvSocketPath))) {
    throw new Error(
      `An mpv/SubMiner session appears to be running (socket ${context.mpvSocketPath}). Close it before syncing, or pass --force.`,
    );
  }
}

// In --json mode every line on stdout is an NDJSON event: human console output
// is silenced and events are written through the original console logger.
function withJsonEvents(deps: SyncFlowDeps): SyncFlowDeps {
  const writeLine = deps.consoleLog;
  return {
    ...deps,
    consoleLog: () => {},
    writeStdout: () => true,
    emitEvent: (event) => writeLine(JSON.stringify(event)),
  };
}

async function runSnapshotMode(
  context: SyncFlowContext,
  dbPath: string,
  deps: SyncFlowDeps,
): Promise<void> {
  await deps.ensureTrackerQuiescent(context, dbPath);
  const outPath = resolveCliPath(context.args.syncSnapshotPath);
  deps.emitEvent({
    type: 'stage',
    stage: 'snapshot-local',
    message: `Snapshotting local database (${dbPath})`,
  });
  deps.createDbSnapshot(dbPath, outPath);
  deps.emitEvent({ type: 'snapshot-created', path: outPath });
  deps.consoleLog(outPath);
}

async function runMergeMode(
  context: SyncFlowContext,
  dbPath: string,
  deps: SyncFlowDeps,
): Promise<void> {
  await deps.ensureTrackerQuiescent(context, dbPath);
  const snapshotPath = resolveCliPath(context.args.syncMergePath);
  deps.emitEvent({
    type: 'stage',
    stage: 'merge-local',
    message: `Merging ${snapshotPath} into the local database`,
  });
  const summary = deps.mergeSnapshotIntoDb(dbPath, snapshotPath);
  deps.emitEvent({ type: 'merge-summary', target: 'local', summary });
  deps.consoleLog(formatMergeSummary(summary));
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

  const runCheck = (checkHost: string, command: string) =>
    deps.runSsh(checkHost, command, {
      batchMode: true,
      connectTimeoutSeconds: 10,
      timeoutMs: 15_000,
    });
  const probe = runCheck(host, 'echo subminer-check-ok');
  const sshOk = probe.status === 0 && probe.stdout.includes('subminer-check-ok');
  if (!sshOk) {
    error = formatRemoteRunError(`Could not reach ${host} over SSH.`, probe);
  } else {
    deps.consoleLog('SSH connection: ok');
    try {
      const flavor = deps.detectRemoteShellFlavor(host, runCheck);
      if (flavor !== 'posix') deps.consoleLog(`Remote platform: Windows (${flavor})`);
      remoteCommand = deps.resolveRemoteSubminerCommand(
        host,
        args.syncRemoteCmd || null,
        flavor,
        runCheck,
      );
      const version = runCheck(host, `${remoteCommand} --version`);
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

// The remote validates --remove-temp against its own tmpdir; this guard only
// keeps garbage output from an earlier failure out of the remote command.
function cleanupRemote(
  host: string,
  remoteCmd: string,
  remoteTmpDir: string,
  quote: (value: string) => string,
  deps: SyncFlowDeps,
): void {
  if (!path.posix.basename(remoteTmpDir).startsWith(SYNC_TEMP_PREFIX)) return;
  deps.runSsh(host, `${remoteCmd} sync --remove-temp ${quote(remoteTmpDir)}`);
}

/**
 * `sync --make-temp` prints the created dir as its last stdout line (a
 * launcher wrapper may log above it). Backslashes are normalized to forward
 * slashes: scp, the remote SubMiner, and Windows itself all accept them, and
 * it keeps the later `${dir}/file` compositions valid on every platform.
 */
function parseRemoteTempDir(stdout: string): string {
  const lines = stdout
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  const candidate = (lines[lines.length - 1] ?? '').replaceAll('\\', '/');
  return path.posix.basename(candidate).startsWith(SYNC_TEMP_PREFIX) ? candidate : '';
}

function parseRemoteMergeSummary(stdout: string): SyncMergeSummary | null {
  for (const line of stdout.split('\n')) {
    const event = parseSyncProgressLine(line);
    if (event?.type === 'merge-summary' && event.target === 'local') return event.summary;
  }
  return null;
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

  const flavor = deps.detectRemoteShellFlavor(host, deps.runSsh);
  const remoteCmd = deps.resolveRemoteSubminerCommand(host, args.syncRemoteCmd || null, flavor);
  const quote = (value: string) => quoteForRemoteShell(flavor, value);
  if (args.logLevel === 'debug') {
    console.error(`Remote subminer command (${flavor}): ${remoteCmd}`);
  }

  const localTmpDir = makeSyncTempDir(deps.mkdtempSync);
  let remoteTmpDir = '';
  let pulledSummary: SyncMergeSummary | null = null;
  try {
    // Signal failures by throwing (not fail(), which exits synchronously and
    // would skip the finally cleanup, leaking temp dirs holding snapshot data).
    // main().catch() reports the message the same way fail() would.
    const mktemp = deps.runSsh(host, `${remoteCmd} sync --make-temp`);
    remoteTmpDir = mktemp.status === 0 ? parseRemoteTempDir(mktemp.stdout) : '';
    if (!remoteTmpDir) {
      throw new Error(
        formatRemoteRunError(`Could not create a temporary directory on ${host}.`, mktemp),
      );
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
        `${remoteCmd} sync --snapshot ${quote(remoteSnapshot)}${forceFlag}`,
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
      deps.consoleLog(formatMergeSummary(summary));
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
        `${remoteCmd} sync --merge ${quote(incomingSnapshot)}${forceFlag}${args.syncJson ? ' --json' : ''}`,
      );
      deps.writeStdout(mergeRun.stdout);
      const remoteSummary = args.syncJson ? parseRemoteMergeSummary(mergeRun.stdout) : null;
      if (remoteSummary) {
        deps.emitEvent({ type: 'merge-summary', target: 'remote', summary: remoteSummary });
      } else if (mergeRun.stdout.trim()) {
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
        cleanupRemote(host, remoteCmd, remoteTmpDir, quote, deps);
      } catch {
        // best effort
      }
    }
  }
}

export async function runSyncFlow(
  context: SyncFlowContext,
  inputDeps: SyncFlowDeps,
): Promise<void> {
  let deps = inputDeps;
  const { args } = context;
  if (args.syncJson) deps = withJsonEvents(deps);

  try {
    if (args.syncMakeTemp) {
      runMakeTempMode(deps);
    } else if (args.syncRemoveTempPath) {
      runRemoveTempMode(context, deps);
    } else {
      const dbPath = resolveSyncDbPath(context, deps);
      if (args.syncCheck) {
        await runCheckMode(context, deps);
      } else if (args.syncSnapshotPath) {
        await runSnapshotMode(context, dbPath, deps);
      } else if (args.syncMergePath) {
        await runMergeMode(context, dbPath, deps);
      } else if (args.syncHost) {
        await runHostSync(context, dbPath, deps);
      } else {
        throw new Error('sync requires a host, --snapshot <file>, or --merge <file>.');
      }
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
}
