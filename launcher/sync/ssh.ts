import { spawnSync } from 'node:child_process';

export interface RemoteRunResult {
  status: number;
  stdout: string;
  stderr: string;
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
 * Run a command on the SSH host. stdin stays attached so interactive prompts
 * can still read from the terminal; stdout/stderr are captured for callers
 * that need actionable remote failure messages.
 */
export function runSsh(host: string, remoteCommand: string): RemoteRunResult {
  assertSafeSshHost(host);
  const result = spawnSync('ssh', [host, remoteCommand], {
    encoding: 'utf8',
    stdio: ['inherit', 'pipe', 'pipe'],
  });
  if (result.error) {
    throw new Error(`Failed to run ssh: ${(result.error as Error).message}`);
  }
  return { status: result.status ?? 1, stdout: result.stdout ?? '', stderr: result.stderr ?? '' };
}

function assertSafeScpEndpoint(endpoint: string): void {
  const colon = endpoint.indexOf(':');
  const slash = endpoint.indexOf('/');
  if (colon <= 0 || (slash !== -1 && slash < colon)) {
    if (endpoint.startsWith('-')) {
      throw new Error(`Refusing to use scp endpoint that looks like an option: ${endpoint}`);
    }
    return;
  }

  const host = endpoint.slice(0, colon);
  const remotePath = endpoint.slice(colon + 1);
  assertSafeSshHost(host);
  if (remotePath.startsWith('-')) {
    throw new Error(`Refusing to use scp remote path that looks like an option: ${remotePath}`);
  }
}

export function runScp(from: string, to: string): void {
  assertSafeScpEndpoint(from);
  assertSafeScpEndpoint(to);
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

const REMOTE_RUNTIME_PATH =
  'PATH="$HOME/.local/bin:$HOME/.bun/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"';

/**
 * Non-interactive SSH shells often miss user-installed launchers and Bun.
 * Probe the launcher under the same deterministic PATH used by sync itself.
 */
export function resolveRemoteSubminerCommand(
  host: string,
  preferred: string | null,
  runRemote: typeof runSsh = runSsh,
): string {
  // Trusted defaults stay unquoted so the remote shell expands `~`; a
  // user-supplied override is shell-quoted to prevent command injection.
  const candidates: Array<{ value: string; invocation: string }> = preferred
    ? [{ value: preferred, invocation: shellQuote(preferred) }]
    : [
        { value: 'subminer', invocation: 'subminer' },
        { value: '~/.local/bin/subminer', invocation: '~/.local/bin/subminer' },
      ];
  for (const candidate of candidates) {
    const command = `${REMOTE_RUNTIME_PATH} ${candidate.invocation}`;
    const probe = runRemote(host, `${command} --help >/dev/null 2>&1`);
    if (probe.status === 0) {
      return command;
    }
  }
  throw new Error(
    preferred
      ? `Remote command not found on ${host}: ${preferred}`
      : `subminer not found on ${host} (tried PATH and ~/.local/bin/subminer). Pass --remote-cmd <path>.`,
  );
}
