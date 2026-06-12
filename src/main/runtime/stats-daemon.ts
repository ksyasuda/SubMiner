import fs from 'node:fs';
import { execFileSync } from 'node:child_process';
import path from 'node:path';

export type BackgroundStatsServerState = {
  pid: number;
  port: number;
  startedAtMs: number;
};

export function readBackgroundStatsServerState(
  statePath: string,
): BackgroundStatsServerState | null {
  try {
    const raw = JSON.parse(
      fs.readFileSync(statePath, 'utf8'),
    ) as Partial<BackgroundStatsServerState>;
    const pid = raw.pid;
    const port = raw.port;
    const startedAtMs = raw.startedAtMs;
    if (
      typeof pid !== 'number' ||
      !Number.isInteger(pid) ||
      pid <= 0 ||
      typeof port !== 'number' ||
      !Number.isInteger(port) ||
      port <= 0 ||
      typeof startedAtMs !== 'number' ||
      !Number.isInteger(startedAtMs) ||
      startedAtMs <= 0
    ) {
      return null;
    }
    return {
      pid,
      port,
      startedAtMs,
    };
  } catch {
    return null;
  }
}

export function writeBackgroundStatsServerState(
  statePath: string,
  state: BackgroundStatsServerState,
): void {
  fs.mkdirSync(path.dirname(statePath), { recursive: true });
  fs.writeFileSync(statePath, JSON.stringify(state, null, 2), 'utf8');
}

export function removeBackgroundStatsServerState(statePath: string): void {
  try {
    fs.rmSync(statePath, { force: true });
  } catch {
    // ignore
  }
}

export function isBackgroundStatsServerProcessAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function readProcessStartedAtMs(pid: number): number | null {
  try {
    if (process.platform === 'win32') {
      const output = execFileSync(
        'powershell.exe',
        [
          '-NoProfile',
          '-Command',
          `(Get-CimInstance Win32_Process -Filter "ProcessId=${pid}").CreationDate.ToUniversalTime().ToString("o")`,
        ],
        { encoding: 'utf8', timeout: 1000 },
      ).trim();
      const parsed = Date.parse(output);
      return Number.isFinite(parsed) ? parsed : null;
    }

    const output = execFileSync('ps', ['-o', 'lstart=', '-p', String(pid)], {
      encoding: 'utf8',
      timeout: 1000,
    }).trim();
    const parsed = Date.parse(output);
    return Number.isFinite(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export function verifyBackgroundStatsServerIdentity(pid: number, startedAtMs: number): boolean {
  const processStartedAtMs = readProcessStartedAtMs(pid);
  if (processStartedAtMs === null) {
    return false;
  }
  const earliestAllowedStateWriteMs = processStartedAtMs - 1_000;
  const latestAllowedStateWriteMs = processStartedAtMs + 60_000;
  return startedAtMs >= earliestAllowedStateWriteMs && startedAtMs <= latestAllowedStateWriteMs;
}

export function resolveBackgroundStatsServerUrl(
  state: Pick<BackgroundStatsServerState, 'port'>,
): string {
  return `http://127.0.0.1:${state.port}`;
}
