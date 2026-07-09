import { spawnSync } from 'node:child_process';

export interface RemoteRunResult {
  status: number;
  stdout: string;
}

/**
 * Run a command on the SSH host. stdin/stderr stay attached to the terminal
 * so key passphrase / password prompts and remote progress output work;
 * stdout is captured for the caller.
 */
export function runSsh(host: string, remoteCommand: string): RemoteRunResult {
  const result = spawnSync('ssh', [host, remoteCommand], {
    encoding: 'utf8',
    stdio: ['inherit', 'pipe', 'inherit'],
  });
  if (result.error) {
    throw new Error(`Failed to run ssh: ${(result.error as Error).message}`);
  }
  return { status: result.status ?? 1, stdout: result.stdout ?? '' };
}

export function runScp(from: string, to: string): void {
  const result = spawnSync('scp', ['-q', from, to], {
    encoding: 'utf8',
    stdio: ['inherit', 'inherit', 'inherit'],
  });
  if (result.error) {
    throw new Error(`Failed to run scp: ${(result.error as Error).message}`);
  }
  if ((result.status ?? 1) !== 0) {
    throw new Error(`scp failed copying ${from} -> ${to}`);
  }
}

export function shellQuote(value: string): string {
  return `'${value.replaceAll("'", `'\\''`)}'`;
}

/**
 * Non-interactive SSH shells often miss ~/.local/bin in PATH, so probe the
 * configured command first and fall back to the default install location.
 */
export function resolveRemoteSubminerCommand(host: string, preferred: string | null): string {
  const candidates = preferred ? [preferred] : ['subminer', '~/.local/bin/subminer'];
  for (const candidate of candidates) {
    const probe = runSsh(host, `command -v ${candidate} >/dev/null 2>&1`);
    if (probe.status === 0) {
      return candidate;
    }
  }
  throw new Error(
    preferred
      ? `Remote command not found on ${host}: ${preferred}`
      : `subminer not found on ${host} (tried PATH and ~/.local/bin/subminer). Pass --remote-cmd <path>.`,
  );
}
