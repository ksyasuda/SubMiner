import { appendLogLine, resolveDefaultLogFilePath as resolveSharedDefaultLogFilePath } from './shared/log-files';

export type LogLevel = 'debug' | 'info' | 'warn' | 'error';
export type LogLevelSource = 'cli' | 'config';

type LogMethod = (message: string, ...meta: unknown[]) => void;

type Logger = {
  debug: LogMethod;
  info: LogMethod;
  warn: LogMethod;
  error: LogMethod;
  child: (childScope: string) => Logger;
};

const LOG_LEVELS: LogLevel[] = ['debug', 'info', 'warn', 'error'];
const LEVEL_PRIORITY: Record<LogLevel, number> = {
  debug: 10,
  info: 20,
  warn: 30,
  error: 40,
};

const DEFAULT_LOG_LEVEL: LogLevel = 'info';

let cliLogLevel: LogLevel | undefined;
let configLogLevel: LogLevel | undefined;

function pad(value: number): string {
  return String(value).padStart(2, '0');
}

function normalizeLogLevel(level: string | undefined): LogLevel | undefined {
  const normalized = (level || '').toLowerCase() as LogLevel;
  return LOG_LEVELS.includes(normalized) ? normalized : undefined;
}

function getEnvLogLevel(): LogLevel | undefined {
  if (!process || !process.env) return undefined;
  return normalizeLogLevel(process.env.SUBMINER_LOG_LEVEL);
}

function formatTimestamp(date: Date): string {
  const year = date.getFullYear();
  const month = pad(date.getMonth() + 1);
  const day = pad(date.getDate());
  const hour = pad(date.getHours());
  const minute = pad(date.getMinutes());
  const second = pad(date.getSeconds());
  return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
}

function resolveMinLevel(): LogLevel {
  const envLevel = getEnvLogLevel();
  if (cliLogLevel) {
    return cliLogLevel;
  }
  if (envLevel) {
    return envLevel;
  }
  if (configLogLevel) {
    return configLogLevel;
  }
  return DEFAULT_LOG_LEVEL;
}

export function setLogLevel(level: string | undefined, source: LogLevelSource = 'cli'): void {
  const normalized = normalizeLogLevel(level);
  if (source === 'cli') {
    cliLogLevel = normalized;
  } else {
    configLogLevel = normalized;
  }
}

function normalizeError(error: Error): { message: string; stack?: string } {
  return {
    message: error.message,
    ...(error.stack ? { stack: error.stack } : {}),
  };
}

function sanitizeMeta(value: unknown): unknown {
  if (value instanceof Error) {
    return normalizeError(value);
  }
  if (typeof value === 'bigint') {
    return value.toString();
  }
  return value;
}

function safeStringify(value: unknown): string {
  if (typeof value === 'string') {
    return value;
  }
  if (
    typeof value === 'number' ||
    typeof value === 'boolean' ||
    typeof value === 'undefined' ||
    value === null
  ) {
    return String(value);
  }
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function resolveLogFilePath(): string {
  const envPath = process.env.SUBMINER_APP_LOG?.trim();
  if (envPath) {
    return envPath;
  }
  return resolveDefaultLogFilePath();
}

export function resolveDefaultLogFilePath(options?: {
  platform?: NodeJS.Platform;
  homeDir?: string;
  appDataDir?: string;
}): string {
  return resolveSharedDefaultLogFilePath('app', options);
}

function appendToLogFile(line: string): void {
  appendLogLine(resolveLogFilePath(), line);
}

function emit(level: LogLevel, scope: string, message: string, meta: unknown[]): void {
  const minLevel = resolveMinLevel();
  if (LEVEL_PRIORITY[level] < LEVEL_PRIORITY[minLevel]) {
    return;
  }

  const timestamp = formatTimestamp(new Date());
  const prefix = `[subminer] - ${timestamp} - ${level.toUpperCase()} - [${scope}] ${message}`;
  const normalizedMeta = meta.map(sanitizeMeta);

  if (normalizedMeta.length === 0) {
    if (level === 'error') {
      console.error(prefix);
    } else if (level === 'warn') {
      console.warn(prefix);
    } else if (level === 'debug') {
      console.debug(prefix);
    } else {
      console.info(prefix);
    }
    appendToLogFile(prefix);
    return;
  }

  const serialized = normalizedMeta.map(safeStringify).join(' ');
  const finalMessage = `${prefix} ${serialized}`;

  if (level === 'error') {
    console.error(finalMessage);
  } else if (level === 'warn') {
    console.warn(finalMessage);
  } else if (level === 'debug') {
    console.debug(finalMessage);
  } else {
    console.info(finalMessage);
  }
  appendToLogFile(finalMessage);
}

export function createLogger(scope: string): Logger {
  const baseScope = scope.trim();
  if (!baseScope) {
    throw new Error('Logger scope is required');
  }

  const logAt = (level: LogLevel): LogMethod => {
    return (message: string, ...meta: unknown[]) => {
      emit(level, baseScope, message, meta);
    };
  };

  return {
    debug: logAt('debug'),
    info: logAt('info'),
    warn: logAt('warn'),
    error: logAt('error'),
    child: (childScope: string): Logger => {
      const normalizedChild = childScope.trim();
      if (!normalizedChild) {
        throw new Error('Child logger scope is required');
      }
      return createLogger(`${baseScope}:${normalizedChild}`);
    },
  };
}
