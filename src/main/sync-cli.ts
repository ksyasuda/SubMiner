import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { parse as parseJsonc } from 'jsonc-parser';
import { resolveConfigDir, resolveConfigFilePath } from '../config/path-resolution';
import { getDefaultMpvSocketPath } from '../shared/mpv-socket-path';
import {
  extractSyncCliTokens,
  parseSyncCliTokens,
  syncCliUsage,
} from '../core/services/stats-sync/cli-args';
import { createDbSnapshot, findLiveStatsDaemonPid } from '../core/services/stats-sync/shared';
import { formatMergeSummary, mergeSnapshotIntoDb } from '../core/services/stats-sync/merge';
import {
  assertSafeSshHost,
  detectRemoteShellFlavor,
  resolveRemoteSubminerCommand,
  runScp,
  runSsh,
} from '../core/services/stats-sync/ssh';
import { openLibsqlSyncDb } from '../core/services/stats-sync/libsql-driver';
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

function resolveAppConfigDir(): string {
  return resolveConfigDir({
    platform: process.platform,
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
}

function resolveTildePath(input: string): string {
  return input.startsWith('~') ? path.join(os.homedir(), input.slice(1)) : path.resolve(input);
}

// Same resolution as the launcher: honor a configured immersionTracking.dbPath
// (raw config read, tolerant of comments), else <configDir>/immersion.sqlite.
function resolveDefaultDbPath(): string {
  const configPath = resolveConfigFilePath({
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
  let configured = '';
  try {
    const data = fs.readFileSync(configPath, 'utf8');
    const parsed = configPath.endsWith('.jsonc') ? parseJsonc(data) : JSON.parse(data);
    const tracking =
      parsed && typeof parsed === 'object' && !Array.isArray(parsed)
        ? (parsed as Record<string, unknown>).immersionTracking
        : null;
    if (tracking && typeof tracking === 'object' && !Array.isArray(tracking)) {
      const dbPath = (tracking as Record<string, unknown>).dbPath;
      if (typeof dbPath === 'string') configured = dbPath.trim();
    }
  } catch {
    // no config or unreadable config → default location
  }
  if (configured) return resolveTildePath(configured);
  return path.join(resolveAppConfigDir(), 'immersion.sqlite');
}

async function canConnectUnixSocket(socketPath: string): Promise<boolean> {
  return await new Promise<boolean>((resolve) => {
    const socket = net.createConnection(socketPath);
    let settled = false;
    const finish = (value: boolean) => {
      if (settled) return;
      settled = true;
      try {
        socket.destroy();
      } catch {
        // ignore
      }
      resolve(value);
    };
    socket.once('connect', () => finish(true));
    socket.once('error', () => finish(false));
    socket.setTimeout(400, () => finish(false));
  });
}

function recordHostSyncResultToDisk(
  host: string,
  status: 'success' | 'error',
  detail: string | null,
): void {
  try {
    const filePath = getSyncHostsPath(resolveAppConfigDir());
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
    createDbSnapshot: (dbPath, outPath) => createDbSnapshot(openLibsqlSyncDb, dbPath, outPath),
    mergeSnapshotIntoDb: (localDbPath, snapshotPath) =>
      mergeSnapshotIntoDb(openLibsqlSyncDb, localDbPath, snapshotPath),
    formatMergeSummary,
    findLiveStatsDaemonPid,
    assertSafeSshHost,
    detectRemoteShellFlavor,
    resolveRemoteSubminerCommand,
    runScp,
    runSsh,
    fail: (message: string): never => {
      throw new Error(message);
    },
    log: (level, configured, message) => {
      if (level === 'debug' && configured !== 'debug') return;
      console.error(message);
    },
    canConnectUnixSocket,
    realpathSync: (candidate) => fs.realpathSync(candidate),
    mkdtempSync: (prefix) => fs.mkdtempSync(prefix),
    rmSync: (target, options) => fs.rmSync(target, options),
    consoleLog: (message) => console.log(message),
    writeStdout: (text) => process.stdout.write(text),
    ensureTrackerQuiescent: async (context, dbPath) =>
      ensureTrackerQuiescentFlow(context, dbPath, deps),
    emitEvent: () => {},
    recordHostSyncResult: recordHostSyncResultToDisk,
    resolveDefaultDbPath,
    resolvePath: resolveTildePath,
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
