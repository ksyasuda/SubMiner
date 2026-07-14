import fs from 'node:fs';
import path from 'node:path';

export type SyncDirection = 'both' | 'push' | 'pull';
export type SyncResultStatus = 'success' | 'error';

export interface SyncHostEntry {
  host: string;
  label: string | null;
  direction: SyncDirection;
  autoSync: boolean;
  addedAtMs: number;
  lastSyncAtMs: number | null;
  lastSyncStatus: SyncResultStatus | null;
  lastSyncDetail: string | null;
}

export interface SyncHostsState {
  version: 1;
  autoSyncIntervalMinutes: number;
  hosts: SyncHostEntry[];
}

export interface SyncHostUpdate {
  host: string;
  label?: string | null;
  direction?: SyncDirection;
  autoSync?: boolean;
}

export interface SyncResultUpdate {
  atMs: number;
  status: SyncResultStatus;
  detail: string | null;
}

const DEFAULT_AUTO_SYNC_INTERVAL_MINUTES = 60;

export function createDefaultSyncHostsState(): SyncHostsState {
  return {
    version: 1,
    autoSyncIntervalMinutes: DEFAULT_AUTO_SYNC_INTERVAL_MINUTES,
    hosts: [],
  };
}

// Hosts are passed as a single spawn argument to ssh, so the only dangerous
// shapes are option injection (leading "-") and multi-token values.
export function isValidSyncHost(host: string): boolean {
  const trimmed = host.trim();
  return trimmed.length > 0 && !trimmed.startsWith('-') && !/\s/.test(trimmed);
}

function asObject(value: unknown): Record<string, unknown> | null {
  return value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function isDirection(value: unknown): value is SyncDirection {
  return value === 'both' || value === 'push' || value === 'pull';
}

function normalizeHostEntry(value: unknown): SyncHostEntry | null {
  const record = asObject(value);
  if (!record) return null;
  const host = typeof record.host === 'string' ? record.host.trim() : '';
  if (!isValidSyncHost(host)) return null;
  return {
    host,
    label: typeof record.label === 'string' && record.label.trim() ? record.label.trim() : null,
    direction: isDirection(record.direction) ? record.direction : 'both',
    autoSync: record.autoSync === true,
    addedAtMs:
      typeof record.addedAtMs === 'number' && Number.isFinite(record.addedAtMs)
        ? record.addedAtMs
        : 0,
    lastSyncAtMs:
      typeof record.lastSyncAtMs === 'number' && Number.isFinite(record.lastSyncAtMs)
        ? record.lastSyncAtMs
        : null,
    lastSyncStatus:
      record.lastSyncStatus === 'success' || record.lastSyncStatus === 'error'
        ? record.lastSyncStatus
        : null,
    lastSyncDetail: typeof record.lastSyncDetail === 'string' ? record.lastSyncDetail : null,
  };
}

export function normalizeSyncHostsState(value: unknown): SyncHostsState {
  const record = asObject(value);
  if (!record || record.version !== 1) return createDefaultSyncHostsState();
  const hosts = Array.isArray(record.hosts)
    ? record.hosts.map(normalizeHostEntry).filter((entry): entry is SyncHostEntry => entry !== null)
    : [];
  const interval = record.autoSyncIntervalMinutes;
  return {
    version: 1,
    autoSyncIntervalMinutes:
      typeof interval === 'number' && Number.isFinite(interval) && interval >= 1
        ? Math.floor(interval)
        : DEFAULT_AUTO_SYNC_INTERVAL_MINUTES,
    hosts,
  };
}

export function upsertSyncHost(
  state: SyncHostsState,
  update: SyncHostUpdate,
  nowMs: number,
): SyncHostsState {
  const host = update.host.trim();
  if (!isValidSyncHost(host)) {
    throw new Error(`Invalid sync host: ${JSON.stringify(update.host)}`);
  }
  const existing = state.hosts.find((entry) => entry.host === host);
  const base: SyncHostEntry = existing ?? {
    host,
    label: null,
    direction: 'both',
    autoSync: false,
    addedAtMs: nowMs,
    lastSyncAtMs: null,
    lastSyncStatus: null,
    lastSyncDetail: null,
  };
  const next: SyncHostEntry = {
    ...base,
    label: update.label !== undefined ? update.label : base.label,
    direction: update.direction ?? base.direction,
    autoSync: update.autoSync ?? base.autoSync,
  };
  const hosts = existing
    ? state.hosts.map((entry) => (entry.host === host ? next : entry))
    : [...state.hosts, next];
  return { ...state, hosts };
}

export function removeSyncHost(state: SyncHostsState, host: string): SyncHostsState {
  return { ...state, hosts: state.hosts.filter((entry) => entry.host !== host.trim()) };
}

export function recordSyncResult(
  state: SyncHostsState,
  host: string,
  result: SyncResultUpdate,
): SyncHostsState {
  const upserted = upsertSyncHost(state, { host }, result.atMs);
  const hosts = upserted.hosts.map((entry) =>
    entry.host === host.trim()
      ? {
          ...entry,
          lastSyncAtMs: result.atMs,
          lastSyncStatus: result.status,
          lastSyncDetail: result.detail,
        }
      : entry,
  );
  return { ...upserted, hosts };
}

export function getSyncHostsPath(configDir: string): string {
  return path.join(configDir, 'sync-hosts.json');
}

export function readSyncHostsState(
  filePath: string,
  deps?: {
    existsSync?: (candidate: string) => boolean;
    readFileSync?: (candidate: string, encoding: BufferEncoding) => string;
  },
): SyncHostsState {
  const existsSync = deps?.existsSync ?? fs.existsSync;
  const readFileSync = deps?.readFileSync ?? fs.readFileSync;
  if (!existsSync(filePath)) return createDefaultSyncHostsState();
  try {
    return normalizeSyncHostsState(JSON.parse(readFileSync(filePath, 'utf8')));
  } catch {
    return createDefaultSyncHostsState();
  }
}

// Written via a sibling temp file + rename so an interrupted write cannot leave
// a truncated sync-hosts.json behind (readSyncHostsState would silently fall
// back to defaults, dropping every configured host).
export function writeSyncHostsState(
  filePath: string,
  state: SyncHostsState,
  deps?: {
    mkdirSync?: (candidate: string, options: { recursive: true }) => void;
    writeFileSync?: (candidate: string, content: string, encoding: BufferEncoding) => void;
    renameSync?: (from: string, to: string) => void;
    rmSync?: (target: string, options: { force: true }) => void;
  },
): void {
  const mkdirSync = deps?.mkdirSync ?? fs.mkdirSync;
  const writeFileSync = deps?.writeFileSync ?? fs.writeFileSync;
  const renameSync = deps?.renameSync ?? fs.renameSync;
  const rmSync = deps?.rmSync ?? fs.rmSync;
  mkdirSync(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.tmp`;
  writeFileSync(tempPath, `${JSON.stringify(state, null, 2)}\n`, 'utf8');
  try {
    renameSync(tempPath, filePath);
  } catch (error) {
    try {
      rmSync(tempPath, { force: true });
    } catch {
      // best effort: the temp file is disposable
    }
    throw error;
  }
}
