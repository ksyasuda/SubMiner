import type { SyncFlowArgs } from './sync-flow';

export const SYNC_CLI_FLAG = '--sync-cli';

export type ParsedSyncCli =
  | { kind: 'help' }
  | { kind: 'version' }
  | { kind: 'run'; args: SyncFlowArgs }
  | { kind: 'error'; message: string };

export function extractSyncCliTokens(argv: readonly string[]): string[] | null {
  const index = argv.indexOf(SYNC_CLI_FLAG);
  if (index === -1) return null;
  return argv.slice(index + 1).filter((token) => token !== SYNC_CLI_FLAG);
}

/**
 * Parse launcher-style sync argv (`sync [host] [--snapshot f] ...`) for the
 * app's --sync-cli mode. Mirrors the launcher's `subminer sync` validation so
 * both entry points accept the same command lines and fail the same way.
 */
export function parseSyncCliTokens(tokens: readonly string[]): ParsedSyncCli {
  if (tokens.includes('--help') || tokens.includes('-h')) return { kind: 'help' };
  if (tokens.includes('--version') || tokens.includes('-V')) return { kind: 'version' };

  const rest = [...tokens];
  if (rest[0] !== 'sync') {
    return {
      kind: 'error',
      message: `Expected a "sync" command after ${SYNC_CLI_FLAG} (e.g. ${SYNC_CLI_FLAG} sync --snapshot <file>).`,
    };
  }
  rest.shift();

  let host = '';
  let snapshot = '';
  let merge = '';
  let push = false;
  let pull = false;
  let check = false;
  let force = false;
  let json = false;
  let makeTemp = false;
  let removeTemp = '';
  let remoteCmd = '';
  let dbPath = '';
  let logLevel = 'warn';

  const valueFlags = new Map<string, (value: string) => void>([
    ['--snapshot', (value) => (snapshot = value.trim())],
    ['--merge', (value) => (merge = value.trim())],
    ['--remove-temp', (value) => (removeTemp = value.trim())],
    ['--remote-cmd', (value) => (remoteCmd = value.trim())],
    ['--db', (value) => (dbPath = value.trim())],
    ['--log-level', (value) => (logLevel = value.trim() || 'warn')],
  ]);

  for (let i = 0; i < rest.length; i += 1) {
    const token = rest[i]!;
    const assignValue = valueFlags.get(token.includes('=') ? token.slice(0, token.indexOf('=')) : token);
    if (assignValue) {
      if (token.includes('=')) {
        assignValue(token.slice(token.indexOf('=') + 1));
        continue;
      }
      const value = rest[i + 1];
      // An option-like token is never a value: `--snapshot --force` must fail
      // loudly instead of writing a snapshot to a file named "--force". Paths
      // that really do start with "-" can still be passed as `--snapshot=-x`.
      if (value === undefined || value.startsWith('-')) {
        return { kind: 'error', message: `Missing value for ${token}.` };
      }
      assignValue(value);
      i += 1;
      continue;
    }
    if (token === '--push') push = true;
    else if (token === '--pull') pull = true;
    else if (token === '--check') check = true;
    else if (token === '--force' || token === '-f') force = true;
    else if (token === '--json') json = true;
    else if (token === '--make-temp') makeTemp = true;
    else if (token.startsWith('-')) {
      return { kind: 'error', message: `Unknown sync option: ${token}` };
    } else if (host) {
      return { kind: 'error', message: `Unexpected extra argument: ${token}` };
    } else {
      host = token.trim();
    }
  }

  if (push && pull) return { kind: 'error', message: 'Sync --push and --pull cannot be combined.' };
  if ((push || pull) && !host) {
    return { kind: 'error', message: 'Sync --push and --pull require a host.' };
  }
  if (check && !host) return { kind: 'error', message: 'Sync --check requires a host.' };
  if (check && (push || pull || snapshot || merge)) {
    return {
      kind: 'error',
      message: 'Sync --check cannot be combined with --push, --pull, --snapshot, or --merge.',
    };
  }
  const modes = [
    Boolean(host),
    Boolean(snapshot),
    Boolean(merge),
    makeTemp,
    Boolean(removeTemp),
  ].filter(Boolean).length;
  if (modes === 0) {
    return { kind: 'error', message: 'Sync requires a host, --snapshot <file>, or --merge <file>.' };
  }
  if (modes > 1) {
    return {
      kind: 'error',
      message: 'Sync host, --snapshot, --merge, --make-temp, and --remove-temp cannot be combined.',
    };
  }

  return {
    kind: 'run',
    args: {
      sync: true,
      syncHost: host,
      syncSnapshotPath: snapshot,
      syncMergePath: merge,
      syncDirection: push ? 'push' : pull ? 'pull' : 'both',
      syncRemoteCmd: remoteCmd,
      syncDbPath: dbPath,
      syncForce: force,
      syncJson: json,
      syncCheck: check,
      syncMakeTemp: makeTemp,
      syncRemoveTempPath: removeTemp,
      logLevel,
    },
  };
}

export function syncCliUsage(): string {
  return [
    'SubMiner sync CLI',
    '',
    `Usage: SubMiner ${SYNC_CLI_FLAG} sync [host] [options]`,
    '',
    'Modes (exactly one):',
    '  <host>               Sync stats with an SSH destination (user@host or ssh alias)',
    '  --snapshot <file>    Write a consistent snapshot of the local stats database',
    '  --merge <file>       Merge a snapshot database file into the local stats database',
    '  --make-temp          Create a sync temp directory and print its path (used over SSH)',
    '  --remove-temp <dir>  Remove a sync temp directory created by --make-temp',
    '',
    'Options:',
    '  --push               Only merge local stats into the SSH host',
    '  --pull               Only merge stats from the SSH host into the local database',
    '  --check              Test the SSH connection and remote SubMiner availability',
    '  --db <file>          Override the local stats database path',
    '  --remote-cmd <cmd>   SubMiner app or launcher command to run on the remote host',
    '  -f, --force          Skip the running-app safety check',
    '  --json               Emit machine-readable NDJSON progress output',
    '  --log-level <level>  Log level',
    '  --help               Show this help',
    '  --version            Show the SubMiner version',
  ].join('\n');
}
