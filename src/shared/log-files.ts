import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export type LogKind = 'app' | 'launcher' | 'mpv';
export type LogRotation = number;
export type LogFileToggles = Record<LogKind, boolean>;

export const DEFAULT_LOG_RETENTION_DAYS = 7;
export const DEFAULT_LOG_MAX_BYTES = 10 * 1024 * 1024;
export const DEFAULT_LOG_ROTATION: LogRotation = DEFAULT_LOG_RETENTION_DAYS;
export const DEFAULT_LOG_FILE_TOGGLES: LogFileToggles = {
  app: true,
  launcher: true,
  mpv: false,
};

const TRUNCATED_MARKER = '[truncated older log content]\n';
const prunedDirectories = new Set<string>();
const NS_PER_MS = 1_000_000n;
const MS_PER_DAY = 86_400_000n;

const LOG_ENABLED_ENV_BY_KIND: Record<LogKind, string> = {
  app: 'SUBMINER_APP_LOG_ENABLED',
  launcher: 'SUBMINER_LAUNCHER_LOG_ENABLED',
  mpv: 'SUBMINER_MPV_LOG_ENABLED',
};

const LOG_PATH_ENV_BY_KIND: Record<LogKind, string> = {
  app: 'SUBMINER_APP_LOG',
  launcher: 'SUBMINER_LAUNCHER_LOG',
  mpv: 'SUBMINER_MPV_LOG',
};

function floorDiv(left: number, right: number): number {
  return Math.floor(left / right);
}

function padDatePart(value: number): string {
  return String(value).padStart(2, '0');
}

function localDateKey(date: Date): string {
  if (Number.isNaN(date.getTime())) {
    throw new RangeError('Invalid time value');
  }
  return `${date.getFullYear()}-${padDatePart(date.getMonth() + 1)}-${padDatePart(date.getDate())}`;
}

function daysFromCivil(year: number, month: number, day: number): bigint {
  const adjustedYear = year - (month <= 2 ? 1 : 0);
  const era = floorDiv(adjustedYear >= 0 ? adjustedYear : adjustedYear - 399, 400);
  const yearOfEra = adjustedYear - era * 400;
  const monthIndex = month + (month > 2 ? -3 : 9);
  const dayOfYear = floorDiv(153 * monthIndex + 2, 5) + day - 1;
  const dayOfEra = yearOfEra * 365 + floorDiv(yearOfEra, 4) - floorDiv(yearOfEra, 100) + dayOfYear;
  return BigInt(era * 146097 + dayOfEra - 719468);
}

function dateToEpochMs(date: Date): bigint {
  const dayCount = daysFromCivil(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
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
    rotation?: unknown;
  },
): string {
  const now = options?.now ?? new Date();
  const suffix = localDateKey(now);
  return path.join(resolveLogBaseDir(options), 'logs', `${kind}-${suffix}.log`);
}

export function normalizeLogRotation(rotation: unknown): LogRotation | undefined {
  const parsed =
    typeof rotation === 'number'
      ? rotation
      : typeof rotation === 'string' && /^\d+$/.test(rotation.trim())
        ? Number(rotation.trim())
        : Number.NaN;
  if (!Number.isInteger(parsed) || parsed <= 0) return undefined;
  return parsed;
}

export function normalizeLogFileEnabled(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') return value;
  if (typeof value !== 'string') return undefined;
  const normalized = value.trim().toLowerCase();
  if (['1', 'true', 'yes', 'on'].includes(normalized)) return true;
  if (['0', 'false', 'no', 'off'].includes(normalized)) return false;
  return undefined;
}

export function isLogFileEnabled(kind: LogKind, env: NodeJS.ProcessEnv = process.env): boolean {
  const configured = normalizeLogFileEnabled(env[LOG_ENABLED_ENV_BY_KIND[kind]]);
  if (configured !== undefined) return configured;
  const explicitPath = env[LOG_PATH_ENV_BY_KIND[kind]]?.trim();
  if (explicitPath) return true;
  return DEFAULT_LOG_FILE_TOGGLES[kind];
}

export function applyLogFileTogglesToEnv(
  files: Partial<LogFileToggles> | undefined,
  env: NodeJS.ProcessEnv = process.env,
): void {
  for (const kind of Object.keys(LOG_ENABLED_ENV_BY_KIND) as LogKind[]) {
    const explicitPath = env[LOG_PATH_ENV_BY_KIND[kind]]?.trim();
    const enabled = files?.[kind] ?? (explicitPath ? true : DEFAULT_LOG_FILE_TOGGLES[kind]);
    env[LOG_ENABLED_ENV_BY_KIND[kind]] = String(enabled);
  }
}

export function getLogRetentionDays(rotation: unknown): number {
  const normalized = normalizeLogRotation(rotation) ?? DEFAULT_LOG_ROTATION;
  return normalized;
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

export function pruneLogDirectoryForPath(logPath: string, rotation?: unknown): void {
  if (!logPath.trim()) return;
  maybePruneLogDirectory(logPath, getLogRetentionDays(rotation));
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
    rotation?: unknown;
    maxBytes?: number;
  },
): void {
  if (!logPath.trim()) return;
  const retentionDays = options?.retentionDays ?? getLogRetentionDays(options?.rotation);
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
