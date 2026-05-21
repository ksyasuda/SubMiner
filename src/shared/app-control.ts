import crypto from 'node:crypto';
import os from 'node:os';
import path from 'node:path';

export const SUBMINER_APP_CONTROL_SOCKET_ENV = 'SUBMINER_APP_CONTROL_SOCKET';

export interface AppControlSocketPathOptions {
  configDir?: string;
  env?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  tmpDir?: string;
}

export interface AppControlRequest {
  argv: string[];
}

export interface AppControlResponse {
  ok: boolean;
  error?: string;
}

function getUserKey(): string {
  if (typeof process.getuid === 'function') {
    return String(process.getuid());
  }
  try {
    const user = os.userInfo();
    if (typeof user.uid === 'number') {
      return String(user.uid);
    }
    if (user.username) {
      return user.username.replace(/[^\w.-]/g, '_');
    }
  } catch {
    // Fall back below.
  }
  return 'user';
}

export function getAppControlSocketPath(options: AppControlSocketPathOptions = {}): string {
  const env = options.env ?? process.env;
  const override = env[SUBMINER_APP_CONTROL_SOCKET_ENV]?.trim();
  if (override) return override;

  const platform = options.platform ?? process.platform;
  const identity = options.configDir?.trim() || 'default';
  const digest = crypto.createHash('sha256').update(identity).digest('hex').slice(0, 16);

  if (platform === 'win32') {
    return `\\\\.\\pipe\\subminer-control-${digest}`;
  }

  return path.join(
    options.tmpDir ?? os.tmpdir(),
    `subminer-control-${getUserKey()}-${digest}.sock`,
  );
}

export function encodeAppControlRequest(argv: string[]): string {
  return `${JSON.stringify({ argv })}\n`;
}

export function encodeAppControlResponse(response: AppControlResponse): string {
  return `${JSON.stringify(response)}\n`;
}

function normalizeArgv(value: unknown): string[] | null {
  if (!Array.isArray(value) || value.length > 128) return null;
  const argv: string[] = [];
  for (const entry of value) {
    if (typeof entry !== 'string' || entry.length > 8192) {
      return null;
    }
    argv.push(entry);
  }
  return argv;
}

export function parseAppControlRequestLine(line: string): AppControlRequest {
  const payload = JSON.parse(line) as { argv?: unknown };
  const argv = normalizeArgv(payload.argv);
  if (!argv) {
    throw new Error('Invalid app-control argv payload');
  }
  return { argv };
}

export function parseAppControlResponseLine(line: string): AppControlResponse {
  const payload = JSON.parse(line) as { ok?: unknown; error?: unknown };
  if (payload.ok === true) {
    return { ok: true };
  }
  return {
    ok: false,
    error: typeof payload.error === 'string' ? payload.error : 'App control command failed',
  };
}
