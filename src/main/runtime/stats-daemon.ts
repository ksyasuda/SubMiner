import fs from 'node:fs';
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

export function resolveBackgroundStatsServerUrl(
  state: Pick<BackgroundStatsServerState, 'port'>,
): string {
  return `http://127.0.0.1:${state.port}`;
}
