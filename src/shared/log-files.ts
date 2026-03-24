import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export type LogKind = 'app' | 'launcher' | 'mpv';

export const DEFAULT_LOG_RETENTION_DAYS = 7;
export const DEFAULT_LOG_MAX_BYTES = 10 * 1024 * 1024;

const TRUNCATED_MARKER = '[truncated older log content]\n';
const prunedDirectories = new Set<string>();

export function resolveLogBaseDir(options?: {
  platform?: NodeJS.Platform;
  homeDir?: string;
  appDataDir?: string;
}): string {
  const platform = options?.platform ?? process.platform;
  const homeDir = options?.homeDir ?? os.homedir();
  return platform === 'win32'
    ? path.join(options?.appDataDir?.trim() || path.join(homeDir, 'AppData', 'Roaming'), 'SubMiner')
    : path.join(homeDir, '.config', 'SubMiner');
}

export function resolveDefaultLogFilePath(
  kind: LogKind = 'app',
  options?: {
    platform?: NodeJS.Platform;
    homeDir?: string;
    appDataDir?: string;
    now?: Date;
  },
): string {
  const date = (options?.now ?? new Date()).toISOString().slice(0, 10);
  return path.join(resolveLogBaseDir(options), 'logs', `${kind}-${date}.log`);
}

export function pruneLogFiles(
  logsDir: string,
  options?: {
    retentionDays?: number;
    now?: Date;
  },
): void {
  const retentionDays = options?.retentionDays ?? DEFAULT_LOG_RETENTION_DAYS;
  if (!Number.isFinite(retentionDays) || retentionDays <= 0) return;

  let entries: string[];
  try {
    entries = fs.readdirSync(logsDir);
  } catch {
    return;
  }

  const cutoffMs = (options?.now ?? new Date()).getTime() - retentionDays * 24 * 60 * 60 * 1000;
  for (const entry of entries) {
    const candidate = path.join(logsDir, entry);
    let stats: fs.Stats;
    try {
      stats = fs.statSync(candidate);
    } catch {
      continue;
    }
    if (!stats.isFile() || !entry.endsWith('.log') || stats.mtimeMs >= cutoffMs) continue;
    try {
      fs.rmSync(candidate, { force: true });
    } catch {
      // ignore cleanup failures
    }
  }
}

function maybePruneLogDirectory(logPath: string, retentionDays: number): void {
  const logsDir = path.dirname(logPath);
  const key = `${logsDir}:${new Date().toISOString().slice(0, 10)}:${retentionDays}`;
  if (prunedDirectories.has(key)) return;
  pruneLogFiles(logsDir, { retentionDays });
  prunedDirectories.add(key);
}

function trimLogFileToMaxBytes(logPath: string, maxBytes: number): void {
  if (!Number.isFinite(maxBytes) || maxBytes <= 0) return;

  let stats: fs.Stats;
  try {
    stats = fs.statSync(logPath);
  } catch {
    return;
  }
  if (stats.size <= maxBytes) return;

  try {
    const buffer = fs.readFileSync(logPath);
    const marker = Buffer.from(TRUNCATED_MARKER, 'utf8');
    const tailBudget = Math.max(0, maxBytes - marker.length);
    const tail =
      tailBudget > 0 ? buffer.subarray(Math.max(0, buffer.length - tailBudget)) : Buffer.alloc(0);
    fs.writeFileSync(logPath, Buffer.concat([marker, tail]));
  } catch {
    // ignore trim failures
  }
}

export function appendLogLine(
  logPath: string,
  line: string,
  options?: {
    retentionDays?: number;
    maxBytes?: number;
  },
): void {
  const retentionDays = options?.retentionDays ?? DEFAULT_LOG_RETENTION_DAYS;
  const maxBytes = options?.maxBytes ?? DEFAULT_LOG_MAX_BYTES;

  try {
    fs.mkdirSync(path.dirname(logPath), { recursive: true });
    maybePruneLogDirectory(logPath, retentionDays);
    fs.appendFileSync(logPath, `${line}\n`, { encoding: 'utf8' });
    trimLogFileToMaxBytes(logPath, maxBytes);
  } catch {
    // never break runtime due to logging sink failures
  }
}
