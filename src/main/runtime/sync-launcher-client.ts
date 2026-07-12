import { spawn as nodeSpawn } from 'node:child_process';
import { parseSyncProgressLine, type SyncProgressEvent } from '../../shared/sync/sync-events';
import { findCommand } from './command-line-launcher-deps';
import { resolveLauncherResourcePath } from './command-line-launcher';

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

// The launcher is a bun script: prefer the PATH-installed `subminer`, fall
// back to running the bundled resource through bun directly.
export function resolveSyncLauncherCommand(
  deps: {
    findCommand?: typeof findCommand;
    resolveResourcePath?: () => string;
    existsSync?: (candidate: string) => boolean;
  } = {},
): SyncLauncherResolution {
  const find = deps.findCommand ?? findCommand;
  const installed = find('subminer', {});
  if (installed) return { command: [installed], error: null };

  const resourcePath = deps.resolveResourcePath
    ? deps.resolveResourcePath()
    : resolveLauncherResourcePath({});
  const bunPath = find('bun', {});
  if (bunPath && resourcePath) return { command: [bunPath, resourcePath], error: null };

  return {
    command: null,
    error:
      'Could not find the subminer launcher. Install the command-line launcher from SubMiner setup (or install bun).',
  };
}

export function runSyncLauncher(options: {
  command: string[];
  args: string[];
  onEvent: (event: SyncProgressEvent) => void;
  onStderr?: (text: string) => void;
  spawn?: SyncLauncherSpawn;
}): SyncLauncherRunHandle {
  const spawn = options.spawn ?? ((command, args) => nodeSpawn(command, args, { stdio: 'pipe' }));
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

  return {
    cancel: () => {
      cancelled = true;
      try {
        child.kill('SIGTERM');
      } catch {
        // process may already be gone
      }
    },
    done,
  };
}
