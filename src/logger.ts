export type LogLevel = "debug" | "info" | "warn" | "error";

type LogMethod = (message: string, ...meta: unknown[]) => void;

type Logger = {
  debug: LogMethod;
  info: LogMethod;
  warn: LogMethod;
  error: LogMethod;
  child: (childScope: string) => Logger;
};

const LOG_LEVELS: LogLevel[] = ["debug", "info", "warn", "error"];
const LEVEL_PRIORITY: Record<LogLevel, number> = {
  debug: 10,
  info: 20,
  warn: 30,
  error: 40,
};

function pad(value: number): string {
  return String(value).padStart(2, "0");
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
  const raw =
    typeof process !== "undefined" && process?.env
      ? process.env.SUBMINER_LOG_LEVEL
      : undefined;
  const normalized = (raw || "").toLowerCase() as LogLevel;
  if (LOG_LEVELS.includes(normalized)) {
    return normalized;
  }
  return "info";
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
  if (typeof value === "bigint") {
    return value.toString();
  }
  return value;
}

function safeStringify(value: unknown): string {
  if (typeof value === "string") {
    return value;
  }
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "undefined" ||
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

function emit(
  level: LogLevel,
  scope: string,
  message: string,
  meta: unknown[],
): void {
  const minLevel = resolveMinLevel();
  if (LEVEL_PRIORITY[level] < LEVEL_PRIORITY[minLevel]) {
    return;
  }

  const timestamp = formatTimestamp(new Date());
  const prefix = `[subminer] - ${timestamp} - ${level.toUpperCase()} - [${scope}] ${message}`;
  const normalizedMeta = meta.map(sanitizeMeta);

  if (normalizedMeta.length === 0) {
    if (level === "error") {
      console.error(prefix);
    } else if (level === "warn") {
      console.warn(prefix);
    } else if (level === "debug") {
      console.debug(prefix);
    } else {
      console.info(prefix);
    }
    return;
  }

  const serialized = normalizedMeta.map(safeStringify).join(" ");
  const finalMessage = `${prefix} ${serialized}`;

  if (level === "error") {
    console.error(finalMessage);
  } else if (level === "warn") {
    console.warn(finalMessage);
  } else if (level === "debug") {
    console.debug(finalMessage);
  } else {
    console.info(finalMessage);
  }
}

export function createLogger(scope: string): Logger {
  const baseScope = scope.trim();
  if (!baseScope) {
    throw new Error("Logger scope is required");
  }

  const logAt = (level: LogLevel): LogMethod => {
    return (message: string, ...meta: unknown[]) => {
      emit(level, baseScope, message, meta);
    };
  };

  return {
    debug: logAt("debug"),
    info: logAt("info"),
    warn: logAt("warn"),
    error: logAt("error"),
    child: (childScope: string): Logger => {
      const normalizedChild = childScope.trim();
      if (!normalizedChild) {
        throw new Error("Child logger scope is required");
      }
      return createLogger(`${baseScope}:${normalizedChild}`);
    },
  };
}
