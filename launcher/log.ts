import fs from 'node:fs';
import path from 'node:path';
import type { LogLevel } from './types.js';
import { DEFAULT_MPV_LOG_FILE } from './types.js';

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
  const envPath = process.env.SUBMINER_MPV_LOG?.trim();
  if (envPath) return envPath;
  return DEFAULT_MPV_LOG_FILE;
}

export function appendToMpvLog(message: string): void {
  const logPath = getMpvLogPath();
  try {
    fs.mkdirSync(path.dirname(logPath), { recursive: true });
    fs.appendFileSync(logPath, `[${new Date().toISOString()}] ${message}\n`, { encoding: 'utf8' });
  } catch {
    // ignore logging failures
  }
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
  appendToMpvLog(`[${level.toUpperCase()}] ${message}`);
}

export function fail(message: string): never {
  process.stderr.write(`${COLORS.red}[ERROR]${COLORS.reset} ${message}\n`);
  appendToMpvLog(`[ERROR] ${message}`);
  process.exit(1);
}
