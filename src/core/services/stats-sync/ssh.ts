import { spawnSync } from 'node:child_process';
import { SYNC_CLI_FLAG } from './cli-args';

export interface RemoteRunResult {
  status: number;
  stdout: string;
  stderr: string;
}

export interface RunSshOptions {
  batchMode?: boolean;
  connectTimeoutSeconds?: number;
  timeoutMs?: number;
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
export function runSsh(
  host: string,
  remoteCommand: string,
  options: RunSshOptions = {},
): RemoteRunResult {
  assertSafeSshHost(host);
  const args: string[] = [];
  if (options.batchMode) args.push('-o', 'BatchMode=yes');
  if (options.connectTimeoutSeconds !== undefined) {
    args.push('-o', `ConnectTimeout=${options.connectTimeoutSeconds}`);
  }
  args.push(host, remoteCommand);
  const result = spawnSync('ssh', args, {
    encoding: 'utf8',
    stdio: ['inherit', 'pipe', 'pipe'],
    timeout: options.timeoutMs,
  });
  if (result.error) {
    throw new Error(`Failed to run ssh: ${(result.error as Error).message}`);
  }
  return { status: result.status ?? 1, stdout: result.stdout ?? '', stderr: result.stderr ?? '' };
}

function assertSafeScpEndpoint(endpoint: string): void {
  if (/^[A-Za-z]:[\\/]/.test(endpoint)) return;
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

/**
 * The shell that Windows OpenSSH hands remote commands to (cmd.exe by
 * default, PowerShell when DefaultShell is changed) — it decides quoting,
 * environment-variable expansion, and which SubMiner install paths to probe.
 */
export type RemoteShellFlavor = 'posix' | 'windows-cmd' | 'windows-powershell';

/**
 * Identify the remote shell with probes that are harmless everywhere:
 * `uname -s` only succeeds on a POSIX shell, `echo %OS%` only expands under
 * cmd.exe, and `echo $env:OS` only expands under PowerShell. Defaults to
 * posix so unreachable/odd hosts fail with the familiar POSIX errors.
 */
export function detectRemoteShellFlavor(
  host: string,
  runRemote: (host: string, remoteCommand: string) => RemoteRunResult,
): RemoteShellFlavor {
  const posixProbe = runRemote(host, 'uname -s');
  if (posixProbe.status === 0 && posixProbe.stdout.trim().length > 0) return 'posix';
  const cmdProbe = runRemote(host, 'echo %OS%');
  if (cmdProbe.status === 0 && cmdProbe.stdout.includes('Windows_NT')) return 'windows-cmd';
  const powershellProbe = runRemote(host, 'echo $env:OS');
  if (powershellProbe.status === 0 && powershellProbe.stdout.includes('Windows_NT')) {
    return 'windows-powershell';
  }
  return 'posix';
}

/**
 * Quote one argument for the detected remote shell. PowerShell expands $(...)
 * and $var inside double quotes, so it gets a single-quoted literal ('' escapes
 * a quote). cmd.exe has no single-quote form and treats ' literally, so it keeps
 * double quotes and rejects values carrying a double quote of their own.
 */
export function quoteForRemoteShell(flavor: RemoteShellFlavor, value: string): string {
  if (flavor === 'posix') return shellQuote(value);
  if (/[\r\n]/.test(value)) {
    throw new Error(`Refusing to quote a value with newlines for a Windows shell: ${value}`);
  }
  if (flavor === 'windows-powershell') {
    return `'${value.replaceAll("'", "''")}'`;
  }
  if (value.includes('"')) {
    throw new Error(`Refusing to quote a value with quotes for a Windows shell: ${value}`);
  }
  if (value.includes('%')) {
    throw new Error(`Refusing to quote a value with percent signs for cmd.exe: ${value}`);
  }
  return `"${value}"`;
}

const REMOTE_RUNTIME_PATH =
  'PATH="$HOME/.local/bin:$HOME/.bun/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"';

// The Electron app answers the same launcher-style `sync ...` argv when
// invoked with --sync-cli (SYNC_CLI_FLAG), so a remote machine only needs the
// app installed; the command-line launcher is one candidate, not a
// requirement. Each candidate is the remote invocation to probe, in order.
function defaultRemoteCandidates(flavor: RemoteShellFlavor): string[] {
  if (flavor === 'windows-cmd' || flavor === 'windows-powershell') {
    // cmd.exe expands %VAR% inside double quotes; PowerShell needs $env: and
    // the & call operator to run a quoted path.
    const launcherShim =
      flavor === 'windows-cmd'
        ? `"%LOCALAPPDATA%\\SubMiner\\bin\\subminer.cmd"`
        : `& "$env:LOCALAPPDATA\\SubMiner\\bin\\subminer.cmd"`;
    const appInstall =
      flavor === 'windows-cmd'
        ? `"%LOCALAPPDATA%\\Programs\\SubMiner\\SubMiner.exe"`
        : `& "$env:LOCALAPPDATA\\Programs\\SubMiner\\SubMiner.exe"`;
    return [
      // Command-line launcher shim on PATH or in its default install dir.
      'subminer',
      launcherShim,
      // The app binary itself in sync-CLI mode (default NSIS install dir).
      `${appInstall} ${SYNC_CLI_FLAG}`,
      `SubMiner ${SYNC_CLI_FLAG}`,
    ];
  }
  return [
    // Command-line launcher (bun script) on PATH or in its default install dir.
    'subminer',
    '~/.local/bin/subminer',
    // The app binary itself in sync-CLI mode.
    `SubMiner ${SYNC_CLI_FLAG}`,
    `/Applications/SubMiner.app/Contents/MacOS/SubMiner ${SYNC_CLI_FLAG}`,
    `~/Applications/SubMiner.app/Contents/MacOS/SubMiner ${SYNC_CLI_FLAG}`,
  ];
}

function preferredCandidates(flavor: RemoteShellFlavor, preferred: string): string[] {
  const quoted = quoteForRemoteShell(flavor, preferred);
  const invocation = flavor === 'windows-powershell' ? `& ${quoted}` : quoted;
  // App binaries also answer plain --help by opening the GUI-oriented help
  // path, so probe the sync-CLI shape first.
  return [`${invocation} ${SYNC_CLI_FLAG}`, invocation];
}

/**
 * Non-interactive POSIX SSH shells often miss user-installed launchers and
 * Bun, so those candidates are probed under the same deterministic PATH sync
 * itself uses; Windows shells get their default install locations instead.
 * Trusted defaults stay unquoted so the remote shell expands `~`/%VAR%; a
 * user-supplied override is quoted to prevent command injection and probed
 * both as an app binary (--sync-cli) and as a launcher.
 */
export function resolveRemoteSubminerCommand(
  host: string,
  preferred: string | null,
  flavor: RemoteShellFlavor = 'posix',
  runRemote: typeof runSsh = runSsh,
): string {
  const candidates = preferred
    ? preferredCandidates(flavor, preferred)
    : defaultRemoteCandidates(flavor);
  for (const candidate of candidates) {
    const command = flavor === 'posix' ? `${REMOTE_RUNTIME_PATH} ${candidate}` : candidate;
    const probe = runRemote(
      host,
      flavor === 'posix' ? `${command} --help >/dev/null 2>&1` : `${command} --help`,
    );
    if (probe.status === 0) {
      return command;
    }
  }
  throw new Error(
    preferred
      ? `Remote command not found on ${host}: ${preferred}`
      : `SubMiner not found on ${host} (tried the subminer launcher and the SubMiner app binary in their default install locations). Pass --remote-cmd <path> pointing at the SubMiner app or launcher.`,
  );
}
