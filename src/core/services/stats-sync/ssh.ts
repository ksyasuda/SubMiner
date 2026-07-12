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

// The Electron app answers the same launcher-style `sync ...` argv when
// invoked with --sync-cli, so a remote machine only needs the app installed;
// the command-line launcher is one candidate, not a requirement.
const APP_SYNC_CLI_FLAG = '--sync-cli';

interface RemoteCandidate {
  invocation: string;
}

function defaultRemoteCandidates(): RemoteCandidate[] {
  return [
    // Command-line launcher (bun script) on PATH or in its default install dir.
    { invocation: 'subminer' },
    { invocation: '~/.local/bin/subminer' },
    // The app binary itself in sync-CLI mode.
    { invocation: `SubMiner ${APP_SYNC_CLI_FLAG}` },
    { invocation: `/Applications/SubMiner.app/Contents/MacOS/SubMiner ${APP_SYNC_CLI_FLAG}` },
    { invocation: `~/Applications/SubMiner.app/Contents/MacOS/SubMiner ${APP_SYNC_CLI_FLAG}` },
  ];
}

/**
 * Non-interactive SSH shells often miss user-installed launchers and Bun.
 * Probe candidates under the same deterministic PATH used by sync itself.
 * Trusted defaults stay unquoted so the remote shell expands `~`; a
 * user-supplied override is shell-quoted to prevent command injection and
 * probed both as an app binary (--sync-cli) and as a launcher.
 */
export function resolveRemoteSubminerCommand(
  host: string,
  preferred: string | null,
  runRemote: typeof runSsh = runSsh,
): string {
  const candidates: RemoteCandidate[] = preferred
    ? [
        // App binaries also answer plain --help by opening the GUI-oriented
        // help path, so probe the sync-CLI shape first.
        { invocation: `${shellQuote(preferred)} ${APP_SYNC_CLI_FLAG}` },
        { invocation: shellQuote(preferred) },
      ]
    : defaultRemoteCandidates();
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
      : `SubMiner not found on ${host} (tried the subminer launcher on PATH and ~/.local/bin, and the SubMiner app binary). Pass --remote-cmd <path> pointing at the SubMiner app or launcher.`,
  );
}
