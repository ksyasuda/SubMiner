import { execFile } from 'node:child_process';
import { getLinuxDesktopEnv, isSupportedWaylandCompositor } from '../../shared/mpv-x11-backend';
import type { PointerPoint } from './linux-overlay-pointer-interaction';

type CommandRunner = (command: string, args: string[]) => Promise<string>;

const XDOTOOL_CURSOR_ARGS = ['getmouselocation', '--shell'] as const;
const CURSOR_POINT_MAX_AGE_MS = 1000;
const COMMAND_FAILURE_RETRY_DELAY_MS = 1000;

function execFileUtf8(command: string, args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(command, args, { encoding: 'utf-8' }, (error, stdout) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(stdout);
    });
  });
}

export function parseXdotoolMouseLocation(raw: string): PointerPoint | null {
  const xMatch = raw.match(/^X=(-?\d+)$/m);
  const yMatch = raw.match(/^Y=(-?\d+)$/m);
  if (!xMatch || !yMatch) return null;

  const x = Number.parseInt(xMatch[1]!, 10);
  const y = Number.parseInt(yMatch[1]!, 10);
  if (!Number.isInteger(x) || !Number.isInteger(y)) return null;
  return { x, y };
}

export function createLinuxX11CursorPointReader(options?: {
  env?: NodeJS.ProcessEnv;
  now?: () => number;
  platform?: NodeJS.Platform;
  runCommand?: CommandRunner;
}) {
  const env = options?.env ?? process.env;
  const now = options?.now ?? (() => Date.now());
  const platform = options?.platform ?? process.platform;
  const runCommand = options?.runCommand ?? execFileUtf8;
  let latest: { point: PointerPoint; updatedAtMs: number } | null = null;
  let inFlight = false;
  let retryAfterMs = 0;

  function isSupported(): boolean {
    if (platform !== 'linux' || !env.DISPLAY?.trim()) return false;
    if (getLinuxDesktopEnv(env).hasWayland && isSupportedWaylandCompositor(env)) return false;
    return true;
  }

  function refresh(): void {
    const nowMs = now();
    if (!isSupported() || inFlight || nowMs < retryAfterMs) return;

    inFlight = true;
    void runCommand('xdotool', [...XDOTOOL_CURSOR_ARGS])
      .then((raw) => {
        const point = parseXdotoolMouseLocation(raw);
        if (!point) {
          retryAfterMs = now() + COMMAND_FAILURE_RETRY_DELAY_MS;
          return;
        }
        latest = { point, updatedAtMs: now() };
        retryAfterMs = 0;
      })
      .catch(() => {
        retryAfterMs = now() + COMMAND_FAILURE_RETRY_DELAY_MS;
      })
      .finally(() => {
        inFlight = false;
      });
  }

  return {
    getCursorScreenPoint(fallback: PointerPoint): PointerPoint {
      refresh();
      if (latest && now() - latest.updatedAtMs <= CURSOR_POINT_MAX_AGE_MS) {
        return latest.point;
      }
      return fallback;
    },
    refresh,
  };
}
