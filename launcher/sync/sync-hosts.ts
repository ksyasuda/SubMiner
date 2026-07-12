import fs from 'node:fs';
import os from 'node:os';
import { resolveConfigDir } from '../../src/config/path-resolution.js';
import {
  getSyncHostsPath,
  readSyncHostsState,
  recordSyncResult,
  writeSyncHostsState,
  type SyncResultStatus,
} from '../../src/shared/sync/sync-hosts-store.js';

export function resolveSyncHostsFilePath(): string {
  const configDir = resolveConfigDir({
    platform: process.platform,
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
  return getSyncHostsPath(configDir);
}

// Best-effort bookkeeping so hosts synced from the CLI show up in the sync UI;
// a failure to persist must never fail the sync itself.
export function recordHostSyncResultToDisk(
  host: string,
  status: SyncResultStatus,
  detail: string | null,
): void {
  try {
    const filePath = resolveSyncHostsFilePath();
    const state = recordSyncResult(readSyncHostsState(filePath), host, {
      atMs: Date.now(),
      status,
      detail,
    });
    writeSyncHostsState(filePath, state);
  } catch {
    // best effort
  }
}
