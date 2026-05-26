import type { LogLevel } from './types.js';
import { getDefaultLauncherLogFile } from './types.js';
import {
  appendLogLine,
  DEFAULT_LOG_ROTATION,
  isLogFileEnabled,
  normalizeLogRotation,
  pruneLogDirectoryForPath,
  resolveDefaultLogFilePath,
  type LogRotation,
} from '../src/shared/log-files.js';

export const COLORS = {
  red: '\x1b[0;31m',
  green: '\x1b[0;32m',
  yellow: '\x1b[0;33m',
  cyan: '\x1b[0;36m',
  reset: '\x1b[0m',
};

export const LOG_PRI: Record<LogLevel, number> = {
  debug: 10,
  info: 20,
  warn: 30,
  error: 40,
};

export function shouldLog(level: LogLevel, configured: LogLevel): boolean {
  return LOG_PRI[level] >= LOG_PRI[configured];
}

export function getMpvLogPath(): string {
  if (!isLogFileEnabled('mpv')) return '';
  const envPath = process.env.SUBMINER_MPV_LOG?.trim();
  const logPath = envPath || resolveDefaultLogFilePath('mpv');
  pruneLogDirectoryForPath(logPath, getLogRotation());
  return logPath;
}

export function getLauncherLogPath(): string {
  if (!isLogFileEnabled('launcher')) return '';
  const envPath = process.env.SUBMINER_LAUNCHER_LOG?.trim();
  if (envPath) return envPath;
  return getDefaultLauncherLogFile();
}

export function getAppLogPath(): string {
  if (!isLogFileEnabled('app')) return '';
  const envPath = process.env.SUBMINER_APP_LOG?.trim();
  if (envPath) return envPath;
  return resolveDefaultLogFilePath('app');
}

function getLogRotation(): LogRotation {
  return normalizeLogRotation(process.env.SUBMINER_LOG_ROTATION) ?? DEFAULT_LOG_ROTATION;
}

function appendTimestampedLog(logPath: string, message: string): void {
  if (!logPath.trim()) return;
  appendLogLine(logPath, `[${new Date().toISOString()}] ${message}`, {
    rotation: getLogRotation(),
  });
}

export function appendToMpvLog(message: string): void {
  appendTimestampedLog(getMpvLogPath(), message);
}

export function appendToLauncherLog(message: string): void {
  appendTimestampedLog(getLauncherLogPath(), message);
}

export function appendToAppLog(message: string): void {
  appendTimestampedLog(getAppLogPath(), message);
}

export function log(level: LogLevel, configured: LogLevel, message: string): void {
  if (!shouldLog(level, configured)) return;
  const color =
    level === 'info'
      ? COLORS.green
      : level === 'warn'
        ? COLORS.yellow
        : level === 'error'
          ? COLORS.red
          : COLORS.cyan;
  process.stdout.write(`${color}[${level.toUpperCase()}]${COLORS.reset} ${message}\n`);
  appendToLauncherLog(`[${level.toUpperCase()}] ${message}`);
}

export function fail(message: string): never {
  process.stderr.write(`${COLORS.red}[ERROR]${COLORS.reset} ${message}\n`);
  appendToLauncherLog(`[ERROR] ${message}`);
  process.exit(1);
}
