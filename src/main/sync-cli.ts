import fs from 'node:fs';
import { getDefaultMpvSocketPath } from '../shared/mpv-socket-path';
import { canConnectSocket } from '../shared/socket-probe';
import { getDefaultConfigDir } from '../shared/setup-state';
import {
  extractSyncCliTokens,
  parseSyncCliTokens,
  syncCliUsage,
} from '../core/services/stats-sync/cli-args';
import { resolveImmersionDbPath } from '../core/services/stats-sync/db-path';
import { createDbSnapshot, findLiveStatsDaemonPid } from '../core/services/stats-sync/shared';
import { mergeSnapshotIntoDb } from '../core/services/stats-sync/merge';
import {
  assertSafeSshHost,
  detectRemoteShellFlavor,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
} from '../core/services/stats-sync/ssh';
import {
  ensureTrackerQuiescentFlow,
  runSyncFlow,
  type SyncFlowDeps,
} from '../core/services/stats-sync/sync-flow';
import {
  recordSyncResult,
  readSyncHostsState,
  writeSyncHostsState,
  getSyncHostsPath,
} from '../shared/sync/sync-hosts-store';

export function shouldHandleSyncCliAtEntry(
  argv: readonly string[],
  env: NodeJS.ProcessEnv,
): boolean {
  if (env.ELECTRON_RUN_AS_NODE === '1') return false;
  return extractSyncCliTokens(argv) !== null;
}

function recordHostSyncResultToDisk(
  host: string,
  status: 'success' | 'error',
  detail: string | null,
): void {
  try {
    const filePath = getSyncHostsPath(getDefaultConfigDir());
    const state = recordSyncResult(readSyncHostsState(filePath), host, {
      atMs: Date.now(),
      status,
      detail,
    });
    writeSyncHostsState(filePath, state);
  } catch {
    // best effort: bookkeeping must never fail the sync itself
  }
}

function buildSyncCliDeps(): SyncFlowDeps {
  const deps: SyncFlowDeps = {
    createDbSnapshot,
    mergeSnapshotIntoDb,
    findLiveStatsDaemonPid,
    assertSafeSshHost,
    detectRemoteShellFlavor,
    resolveRemoteSubminerCommand,
    runScp,
    runSsh,
    canConnectUnixSocket: canConnectSocket,
    realpathSync: (candidate) => fs.realpathSync(candidate),
    mkdtempSync: (prefix) => fs.mkdtempSync(prefix),
    rmSync: (target, options) => fs.rmSync(target, options),
    consoleLog: (message) => console.log(message),
    writeStdout: (text) => process.stdout.write(text),
    ensureTrackerQuiescent: async (context, dbPath) =>
      ensureTrackerQuiescentFlow(context, dbPath, deps),
    emitEvent: () => {},
    recordHostSyncResult: recordHostSyncResultToDisk,
    resolveDefaultDbPath: resolveImmersionDbPath,
  };
  return deps;
}

/**
 * Headless sync entry for the Electron app: answers the same launcher-style
 * `sync ...` argv as `subminer sync`, so a machine only needs the app
 * installed to participate in stats sync (locally or as an SSH remote). Runs
 * before any window/display initialization and never touches Electron APIs.
 */
export async function runSyncCliFromProcess(
  argv: readonly string[],
  appVersion: string,
): Promise<number> {
  const tokens = extractSyncCliTokens(argv);
  const parsed = parseSyncCliTokens(tokens ?? []);
  if (parsed.kind === 'help') {
    console.log(syncCliUsage());
    return 0;
  }
  if (parsed.kind === 'version') {
    console.log(`SubMiner ${appVersion}`);
    return 0;
  }
  if (parsed.kind === 'error') {
    console.error(parsed.message);
    console.error(`\n${syncCliUsage()}`);
    return 2;
  }

  const context = {
    args: parsed.args,
    mpvSocketPath: getDefaultMpvSocketPath(process.platform),
  };
  try {
    await runSyncFlow(context, buildSyncCliDeps());
    return 0;
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    return 1;
  }
}

export async function handleSyncCliAtEntry(
  argv: readonly string[],
  env: NodeJS.ProcessEnv,
  appVersion: string,
  deps: {
    run: typeof runSyncCliFromProcess;
    exit: (code: number) => void;
  } = {
    run: runSyncCliFromProcess,
    exit: (code) => process.exit(code),
  },
): Promise<boolean> {
  if (!shouldHandleSyncCliAtEntry(argv, env)) return false;
  const exitCode = await deps.run(argv, appVersion);
  // This path runs before app readiness and must not call app.exit(): on Linux
  // that can throw and make the entrypoint fall back to full GUI startup.
  deps.exit(exitCode);
  return true;
}
