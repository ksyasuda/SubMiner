import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { assertSafeSshHost, runScp, runSsh, shellQuote, type RemoteShellFlavor } from './ssh';

const RSYNC_OPTIONS = ['--compress', '--checksum'];

export function runRsync(args: string[], timeoutMs = 30 * 60_000) {
  return spawnSync('rsync', ['--rsh=ssh', ...args], {
    encoding: 'utf8',
    stdio: ['inherit', 'pipe', 'pipe'],
    timeout: timeoutMs,
    killSignal: 'SIGKILL',
    // Quote remote paths ourselves for both modern rsync and macOS openrsync.
    env: { ...process.env, RSYNC_OLD_ARGS: '1' },
  });
}

interface TransferDeps {
  platform: NodeJS.Platform;
  runRsync: typeof runRsync;
  runSsh: typeof runSsh;
  runScp: typeof runScp;
}

export interface SnapshotTransfer {
  kind: 'rsync' | 'scp';
  copy: (request: {
    direction: 'download' | 'upload';
    localPath: string;
    remotePath: string;
  }) => void;
}

/**
 * For rsync, both paths name snapshot.sqlite, with the destination inside an
 * incoming/ directory seeded from the transfer cache. rsync creates incoming/
 * when there is no cached basis and verifies the reconstructed file.
 * Missing rsync and Windows endpoints use compressed scp with ordinary paths.
 */
export function createSnapshotTransfer(
  host: string,
  flavor: RemoteShellFlavor,
  deps: TransferDeps = { platform: process.platform, runRsync, runSsh, runScp },
): SnapshotTransfer {
  assertSafeSshHost(host);
  const canUseRsync =
    deps.platform !== 'win32' &&
    flavor === 'posix' &&
    deps.runRsync([...RSYNC_OPTIONS, '--version']).status === 0 &&
    deps.runSsh(host, `rsync ${RSYNC_OPTIONS.join(' ')} --version`, {
      batchMode: true,
      connectTimeoutSeconds: 10,
      timeoutMs: 15_000,
    }).status === 0;

  return {
    kind: canUseRsync ? 'rsync' : 'scp',
    copy: ({ direction, localPath, remotePath }) => {
      // A directory destination lets rsync create incoming/ on either end,
      // including peers running older SubMiner versions.
      const local =
        canUseRsync && direction === 'download' ? `${path.dirname(localPath)}/` : localPath;
      const remoteTarget =
        canUseRsync && direction === 'upload' ? `${path.posix.dirname(remotePath)}/` : remotePath;
      const remote = `${host}:${canUseRsync ? shellQuote(remoteTarget) : remoteTarget}`;
      const [from, to] =
        direction === 'download' ? ([remote, local] as const) : ([local, remote] as const);
      if (!canUseRsync) {
        deps.runScp(from, to);
        return;
      }
      // --checksum prevents a same-size, same-mtime snapshot being skipped.
      // Without --inplace, rsync replaces the staged basis only after the
      // reconstructed file passes its transfer checksum.
      const result = deps.runRsync([...RSYNC_OPTIONS, '--quiet', '--', from, to]);
      if (result.error && 'code' in result.error && result.error.code === 'ETIMEDOUT') {
        throw new Error(`rsync ${direction} timed out for ${host}`);
      }
      if (result.error) throw new Error(`Failed to run rsync: ${result.error.message}`);
      if (result.status !== 0) {
        throw new Error(`rsync ${direction} failed for ${host}: ${result.stderr.trim()}`);
      }
    },
  };
}
