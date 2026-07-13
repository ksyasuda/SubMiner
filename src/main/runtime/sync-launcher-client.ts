import { spawn as nodeSpawn } from 'node:child_process';
import { parseSyncProgressLine, type SyncProgressEvent } from '../../shared/sync/sync-events';
import { SYNC_CLI_FLAG } from '../../core/services/stats-sync/cli-args';

/** How long a cancelled sync child gets to exit on SIGTERM before SIGKILL. */
const CANCEL_GRACE_MS = 5000;

export interface SyncLauncherChildLike {
  stdout: { on(event: 'data', listener: (chunk: Buffer | string) => void): unknown } | null;
  stderr: { on(event: 'data', listener: (chunk: Buffer | string) => void): unknown } | null;
  on(event: 'close', listener: (code: number | null, signal: string | null) => void): unknown;
  on(event: 'exit', listener: (code: number | null, signal: string | null) => void): unknown;
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
  timeoutMs?: number;
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
  let terminationError: string | null = null;
  let settleAfterTerminalEvent: (() => void) | null = null;

  child.stdout?.on('data', (chunk) => {
    stdoutBuffer += chunk.toString();
    let newlineIndex = stdoutBuffer.indexOf('\n');
    while (newlineIndex !== -1) {
      const line = stdoutBuffer.slice(0, newlineIndex);
      stdoutBuffer = stdoutBuffer.slice(newlineIndex + 1);
      const event = parseSyncProgressLine(line);
      if (event) {
        if (event.type === 'result') {
          resultEvent = event;
          settleAfterTerminalEvent?.();
        }
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

  let killTimer: ReturnType<typeof setTimeout> | null = null;
  let operationTimer: ReturnType<typeof setTimeout> | null = null;
  const clearTimers = (): void => {
    if (killTimer !== null) clearTimeout(killTimer);
    if (operationTimer !== null) clearTimeout(operationTimer);
    killTimer = null;
    operationTimer = null;
  };

  const done = new Promise<SyncLauncherRunResult>((resolve) => {
    let settled = false;
    let exitObserved = false;
    let exitCode: number | null = null;
    const settle = (result: SyncLauncherRunResult): void => {
      if (settled) return;
      settled = true;
      clearTimers();
      resolve(result);
    };
    child.on('error', (error) => {
      settle({ ok: false, error: error.message });
    });
    const settleFromExit = (code: number | null) => {
      if (terminationError) {
        settle({ ok: false, error: terminationError });
        return;
      }
      if (code === 0) {
        settle({ ok: true, error: null });
        return;
      }
      const error =
        resultEvent?.error ?? (stderrTail.trim() || `Launcher exited with code ${code ?? 'null'}.`);
      settle({ ok: false, error });
    };
    settleAfterTerminalEvent = () => {
      if (exitObserved) settleFromExit(exitCode);
    };
    // Electron descendants can retain inherited stdio pipes after the main
    // child exits, delaying `close` indefinitely. Once terminal NDJSON has
    // arrived, `exit` is authoritative; keep `close` as the fallback.
    child.on('exit', (code) => {
      exitObserved = true;
      exitCode = code;
      if (resultEvent) settleFromExit(code);
    });
    child.on('close', settleFromExit);
  });

  // A sync child blocked on an ssh password prompt may ignore SIGTERM, so
  // escalate to SIGKILL after a grace period.
  const terminate = (error: string): void => {
    if (terminationError) return;
    terminationError = error;
    killTimer = setTimeout(() => {
      killTimer = null;
      try {
        child.kill('SIGKILL');
      } catch {
        // process may already be gone
      }
    }, CANCEL_GRACE_MS);
    killTimer.unref?.();
    try {
      child.kill('SIGTERM');
    } catch {
      // process may already be gone
    }
  };

  if (options.timeoutMs !== undefined) {
    operationTimer = setTimeout(() => terminate('Sync operation timed out.'), options.timeoutMs);
    operationTimer.unref?.();
  }

  return {
    cancel: () => terminate('Sync cancelled.'),
    done,
  };
}
