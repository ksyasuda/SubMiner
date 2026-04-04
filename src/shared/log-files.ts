import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export type LogKind = 'app' | 'launcher' | 'mpv';

export const DEFAULT_LOG_RETENTION_DAYS = 7;
export const DEFAULT_LOG_MAX_BYTES = 10 * 1024 * 1024;

const TRUNCATED_MARKER = '[truncated older log content]\n';
const prunedDirectories = new Set<string>();
const NS_PER_MS = 1_000_000n;
const MS_PER_DAY = 86_400_000n;

function floorDiv(left: number, right: number): number {
  return Math.floor(left / right);
}

function daysFromCivil(year: number, month: number, day: number): bigint {
  const adjustedYear = year - (month <= 2 ? 1 : 0);
  const era = floorDiv(adjustedYear >= 0 ? adjustedYear : adjustedYear - 399, 400);
  const yearOfEra = adjustedYear - era * 400;
  const monthIndex = month + (month > 2 ? -3 : 9);
  const dayOfYear = floorDiv(153 * monthIndex + 2, 5) + day - 1;
  const dayOfEra =
    yearOfEra * 365 + floorDiv(yearOfEra, 4) - floorDiv(yearOfEra, 100) + dayOfYear;
  return BigInt(era * 146097 + dayOfEra - 719468);
}

function dateToEpochMs(date: Date): bigint {
  const dayCount = daysFromCivil(
    date.getUTCFullYear(),
    date.getUTCMonth() + 1,
    date.getUTCDate(),
  );
  const timeOfDayMs = BigInt(
    ((date.getUTCHours() * 60 + date.getUTCMinutes()) * 60 + date.getUTCSeconds()) * 1000 +
      date.getUTCMilliseconds(),
  );
  return dayCount * MS_PER_DAY + timeOfDayMs;
}

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

  const cutoffDate = new Date(options?.now ?? new Date());
  cutoffDate.setUTCDate(cutoffDate.getUTCDate() - retentionDays);
  const cutoffNs = dateToEpochMs(cutoffDate) * NS_PER_MS;
  for (const entry of entries) {
    const candidate = path.join(logsDir, entry);
    let stats: fs.BigIntStats;
    try {
      stats = fs.statSync(candidate, { bigint: true });
    } catch {
      continue;
    }
    if (!stats.isFile() || !entry.endsWith('.log') || stats.mtimeNs >= cutoffNs) {
      continue;
    }
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
