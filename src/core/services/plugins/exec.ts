import { type ChildProcess, spawn } from 'node:child_process';

import { buildSystemBinaryEnv } from '../../utils/system-binary-env';
import {
  buildWindowsShimInvocation,
  isWindowsCommandShim,
  resolveWindowsCommand,
} from '../../utils/windows-command-shim';
import type { PluginExecOptions, PluginExecResult } from './types';

const MAX_CAPTURED_OUTPUT_CHARS = 4 * 1024 * 1024;
const DEFAULT_KILL_GRACE_MS = 3_000;

function appendCapped(current: string, chunk: string): string {
  if (current.length >= MAX_CAPTURED_OUTPUT_CHARS) return current;
  return (current + chunk).slice(0, MAX_CAPTURED_OUTPUT_CHARS);
}

interface SpawnTarget {
  file: string;
  args: string[];
  viaShell: boolean;
}

/**
 * Windows batch shims (npm installs `codex` as `codex.cmd`) only run through cmd.exe. Native
 * binaries, and every command on Linux and macOS, are spawned directly.
 */
export function resolvePluginSpawnTarget(
  command: string,
  args: string[],
  env: NodeJS.ProcessEnv,
  platform: NodeJS.Platform = process.platform,
): SpawnTarget {
  if (platform !== 'win32') return { file: command, args, viaShell: false };
  const resolved = resolveWindowsCommand(command, env);
  if (!isWindowsCommandShim(resolved)) return { file: resolved, args, viaShell: false };
  return { ...buildWindowsShimInvocation(resolved, args, env), viaShell: true };
}

// Killing cmd.exe alone would leave the shim's own child running, so end the whole tree.
function killWindowsProcessTree(child: ChildProcess): void {
  if (child.pid === undefined) return;
  spawn('taskkill', ['/pid', String(child.pid), '/T', '/F'], {
    stdio: 'ignore',
    windowsHide: true,
  }).on('error', () => child.kill());
}

/**
 * Runs `command` (no shell, except cmd.exe for Windows batch shims) and collects its output. Resolves with the exit code on any exit,
 * including non-zero; rejects when the binary cannot start, the timeout passes, or `signal` aborts.
 * A timed-out or aborted child gets SIGTERM, then SIGKILL after `killGraceMs` (a shim's
 * process tree is ended with taskkill instead).
 */
export function execPluginCommand(
  command: string,
  args: string[],
  options: PluginExecOptions = {},
  killGraceMs = DEFAULT_KILL_GRACE_MS,
): Promise<PluginExecResult> {
  return new Promise((resolve, reject) => {
    if (options.signal?.aborted) {
      reject(new DOMException('Command aborted.', 'AbortError'));
      return;
    }

    const env = { ...buildSystemBinaryEnv(), ...options.env };
    const target = resolvePluginSpawnTarget(command, args, env);
    const child = spawn(target.file, target.args, {
      cwd: options.cwd,
      env,
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
      windowsVerbatimArguments: target.viaShell,
    });

    let stdout = '';
    let stderr = '';
    let settled = false;
    let timeout: ReturnType<typeof setTimeout> | null = null;

    const settle = (finish: () => void): void => {
      if (settled) return;
      settled = true;
      if (timeout) clearTimeout(timeout);
      options.signal?.removeEventListener('abort', onAbort);
      finish();
    };

    // Rejects right away, then escalates to SIGKILL if the child ignores SIGTERM.
    const killWith = (error: Error): void => {
      settle(() => reject(error));
      if (target.viaShell) {
        killWindowsProcessTree(child);
        return;
      }
      child.kill('SIGTERM');
      const forceKill = setTimeout(() => {
        if (child.exitCode === null && child.signalCode === null) child.kill('SIGKILL');
      }, killGraceMs);
      forceKill.unref();
      child.once('exit', () => clearTimeout(forceKill));
    };

    function onAbort(): void {
      killWith(new DOMException('Command aborted.', 'AbortError'));
    }

    options.signal?.addEventListener('abort', onAbort, { once: true });
    if (options.timeoutMs !== undefined && options.timeoutMs > 0) {
      timeout = setTimeout(
        () => killWith(new Error(`${command} timed out after ${options.timeoutMs} ms.`)),
        options.timeoutMs,
      );
    }

    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk: string) => {
      stdout = appendCapped(stdout, chunk);
    });
    child.stderr.on('data', (chunk: string) => {
      stderr = appendCapped(stderr, chunk);
    });
    // A child that exits without reading stdin raises EPIPE here; the exit code tells the story.
    child.stdin.on('error', () => {});

    child.on('error', (error) => settle(() => reject(error)));
    child.on('close', (exitCode) => settle(() => resolve({ exitCode, stdout, stderr })));

    child.stdin.end(options.stdin ?? '');
  });
}
