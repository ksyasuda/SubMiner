import fs from 'node:fs';
import { fail, log } from '../log.js';
import type { LogLevel } from '../types.js';
import { resolveImmersionDbPath } from '../history-db.js';
import {
  createDbSnapshot,
  findLiveStatsDaemonPid,
  formatMergeSummary,
  mergeSnapshotIntoDb,
} from '../sync/sync-db.js';
import {
  assertSafeSshHost,
  detectRemoteShellFlavor,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
} from '../sync/ssh.js';
import { resolvePathMaybe } from '../util.js';
import { canConnectUnixSocket } from '../mpv.js';
import { recordHostSyncResultToDisk } from '../sync/sync-hosts.js';
import {
  ensureTrackerQuiescentFlow,
  runSyncFlow,
  runCheckMode as runCheckModeFlow,
  runHostSync as runHostSyncFlow,
  runMergeMode as runMergeModeFlow,
  runSnapshotMode as runSnapshotModeFlow,
  type SyncFlowContext,
  type SyncFlowDeps,
} from '../../src/core/services/stats-sync/sync-flow.js';
import type { LauncherCommandContext } from './context.js';

// The sync flow itself lives in src/core/services/stats-sync (shared with the
// app's --sync-cli mode); this module binds it to the launcher's bun:sqlite
// engine, logger, and config paths.
export type SyncCommandDeps = SyncFlowDeps;

const defaultSyncCommandDeps: SyncCommandDeps = {
  createDbSnapshot,
  mergeSnapshotIntoDb,
  formatMergeSummary,
  findLiveStatsDaemonPid,
  assertSafeSshHost,
  detectRemoteShellFlavor,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
  fail,
  log: (level, configured, message) => log(level as LogLevel, configured as LogLevel, message),
  canConnectUnixSocket,
  realpathSync: fs.realpathSync,
  mkdtempSync: fs.mkdtempSync,
  rmSync: fs.rmSync,
  consoleLog: console.log,
  writeStdout: process.stdout.write.bind(process.stdout),
  ensureTrackerQuiescent: async (context, dbPath) => ensureTrackerQuiescent(context, dbPath),
  emitEvent: () => {},
  recordHostSyncResult: recordHostSyncResultToDisk,
  resolveDefaultDbPath: resolveImmersionDbPath,
  resolvePath: resolvePathMaybe,
};

function resolveSyncCommandDeps(inputDeps: Partial<SyncCommandDeps> = {}): SyncCommandDeps {
  return { ...defaultSyncCommandDeps, ...inputDeps };
}

export async function ensureTrackerQuiescent(
  context: SyncFlowContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<void> {
  await ensureTrackerQuiescentFlow(context, dbPath, resolveSyncCommandDeps(inputDeps));
}

export function runSnapshotMode(
  context: LauncherCommandContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): void {
  runSnapshotModeFlow(context, dbPath, resolveSyncCommandDeps(inputDeps));
}

export async function runMergeMode(
  context: LauncherCommandContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<void> {
  await runMergeModeFlow(context, dbPath, resolveSyncCommandDeps(inputDeps));
}

export async function runCheckMode(
  context: LauncherCommandContext,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<void> {
  await runCheckModeFlow(context, resolveSyncCommandDeps(inputDeps));
}

export async function runHostSync(
  context: LauncherCommandContext,
  dbPath: string,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<void> {
  await runHostSyncFlow(context, dbPath, resolveSyncCommandDeps(inputDeps));
}

export async function runSyncCommand(
  context: LauncherCommandContext,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<boolean> {
  return runSyncFlow(context, resolveSyncCommandDeps(inputDeps));
}
