import { spawnSync } from 'node:child_process';

export interface RemoteRunResult {
  status: number;
  stdout: string;
}

/**
 * ssh/scp have no `--` terminator for the destination, so a host that starts
 * with `-` (e.g. `-oProxyCommand=...`) is parsed as an option. Reject those
 * before spawning.
 */
export function assertSafeSshHost(host: string): void {
  if (host.startsWith('-')) {
    throw new Error(`Refusing to use SSH host that looks like an option: ${host}`);
  }
}

/**
 * Run a command on the SSH host. stdin/stderr stay attached to the terminal
 * so key passphrase / password prompts and remote progress output work;
 * stdout is captured for the caller.
 */
export function runSsh(host: string, remoteCommand: string): RemoteRunResult {
  assertSafeSshHost(host);
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
  // Trusted defaults stay unquoted so the remote shell expands `~`; the
  // user-supplied override is shell-quoted to prevent command injection.
  const candidates: Array<{ value: string; probe: string }> = preferred
    ? [{ value: preferred, probe: shellQuote(preferred) }]
    : [
        { value: 'subminer', probe: 'subminer' },
        { value: '~/.local/bin/subminer', probe: '~/.local/bin/subminer' },
      ];
  for (const candidate of candidates) {
    const probe = runSsh(host, `command -v ${candidate.probe} >/dev/null 2>&1`);
    if (probe.status === 0) {
      return candidate.value;
    }
  }
  throw new Error(
    preferred
      ? `Remote command not found on ${host}: ${preferred}`
      : `subminer not found on ${host} (tried PATH and ~/.local/bin/subminer). Pass --remote-cmd <path>.`,
  );
}
