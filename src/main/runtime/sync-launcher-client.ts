import { spawn as nodeSpawn } from 'node:child_process';
import { parseSyncProgressLine, type SyncProgressEvent } from '../../shared/sync/sync-events';
import { SYNC_CLI_FLAG } from '../../core/services/stats-sync/cli-args';

/** How long a cancelled sync child gets to exit on SIGTERM before SIGKILL. */
const CANCEL_GRACE_MS = 5000;

export interface SyncLauncherChildLike {
  stdout: { on(event: 'data', listener: (chunk: Buffer | string) => void): unknown } | null;
  stderr: { on(event: 'data', listener: (chunk: Buffer | string) => void): unknown } | null;
  on(event: 'close', listener: (code: number | null, signal: string | null) => void): unknown;
  on(event: 'error', listener: (error: Error) => void): unknown;
  kill(signal?: NodeJS.Signals): boolean;
}

export type SyncLauncherSpawn = (command: string, args: string[]) => SyncLauncherChildLike;

export interface SyncLauncherRunResult {
  ok: boolean;
  error: string | null;
}

export interface SyncLauncherRunHandle {
  cancel: () => void;
  done: Promise<SyncLauncherRunResult>;
}

export interface SyncLauncherResolution {
  command: string[] | null;
  error: string | null;
}

// Sync runs in a child copy of this app in headless --sync-cli mode: same
// engine and NDJSON protocol as `subminer sync --json`, with no dependency on
// bun or an installed command-line launcher. In dev runs process.execPath is
// a bare electron binary, so the app path is passed as its entry argument.
export function resolveSyncLauncherCommand(
  deps: {
    execPath?: string;
    appPath?: string | null;
  } = {},
): SyncLauncherResolution {
  const execPath = deps.execPath ?? process.execPath;
  const appPath = deps.appPath ?? null;
  return {
    command: appPath ? [execPath, appPath, SYNC_CLI_FLAG] : [execPath, SYNC_CLI_FLAG],
    error: null,
  };
}

export function runSyncLauncher(options: {
  command: string[];
  args: string[];
  onEvent: (event: SyncProgressEvent) => void;
  onStderr?: (text: string) => void;
  spawn?: SyncLauncherSpawn;
}): SyncLauncherRunHandle {
  const spawn =
    options.spawn ??
    ((command, args) => {
      // The child must boot as a full Electron app (its entry handles
      // --sync-cli); a leaked ELECTRON_RUN_AS_NODE would turn it into node.
      const env = { ...process.env };
      delete env.ELECTRON_RUN_AS_NODE;
      return nodeSpawn(command, args, { stdio: 'pipe', env });
    });
  const [executable, ...prefixArgs] = options.command;
  const child = spawn(executable!, [...prefixArgs, ...options.args]);

  let stdoutBuffer = '';
  let stderrTail = '';
  let resultEvent: Extract<SyncProgressEvent, { type: 'result' }> | null = null;
  let cancelled = false;

  child.stdout?.on('data', (chunk) => {
    stdoutBuffer += chunk.toString();
    let newlineIndex = stdoutBuffer.indexOf('\n');
    while (newlineIndex !== -1) {
      const line = stdoutBuffer.slice(0, newlineIndex);
      stdoutBuffer = stdoutBuffer.slice(newlineIndex + 1);
      const event = parseSyncProgressLine(line);
      if (event) {
        if (event.type === 'result') resultEvent = event;
        options.onEvent(event);
      }
      newlineIndex = stdoutBuffer.indexOf('\n');
    }
  });
  child.stderr?.on('data', (chunk) => {
    const text = chunk.toString();
    stderrTail = `${stderrTail}${text}`.slice(-4000);
    options.onStderr?.(text);
  });

  const done = new Promise<SyncLauncherRunResult>((resolve) => {
    let settled = false;
    const settle = (result: SyncLauncherRunResult): void => {
      if (settled) return;
      settled = true;
      resolve(result);
    };
    child.on('error', (error) => {
      settle({ ok: false, error: error.message });
    });
    child.on('close', (code) => {
      if (cancelled) {
        settle({ ok: false, error: 'Sync cancelled.' });
        return;
      }
      if (code === 0) {
        settle({ ok: true, error: null });
        return;
      }
      const error =
        resultEvent?.error ?? (stderrTail.trim() || `Launcher exited with code ${code ?? 'null'}.`);
      settle({ ok: false, error });
    });
  });

  // A sync child blocked on an ssh password prompt ignores SIGTERM, so escalate
  // to SIGKILL if it is still alive after a grace period. The timer is cleared
  // on close so a cancelled run cannot hold the process open.
  let killTimer: ReturnType<typeof setTimeout> | null = null;
  const clearKillTimer = (): void => {
    if (killTimer === null) return;
    clearTimeout(killTimer);
    killTimer = null;
  };
  child.on('close', clearKillTimer);

  return {
    cancel: () => {
      if (cancelled) return;
      cancelled = true;
      try {
        child.kill('SIGTERM');
      } catch {
        // process may already be gone
      }
      killTimer = setTimeout(() => {
        killTimer = null;
        try {
          child.kill('SIGKILL');
        } catch {
          // process may already be gone
        }
      }, CANCEL_GRACE_MS);
      killTimer.unref?.();
    },
    done,
  };
}
