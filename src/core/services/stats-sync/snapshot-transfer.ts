import { spawnSync } from 'node:child_process';
import { assertSafeSshHost, runScp, runSsh, shellQuote, type RemoteShellFlavor } from './ssh';

const RSYNC_OPTIONS = ['--compress', '--checksum', '--fuzzy'];

function runRsync(args: string[]) {
  return spawnSync('rsync', args, {
    encoding: 'utf8',
    stdio: ['inherit', 'pipe', 'pipe'],
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
 * Reuse the receiving machine's snapshot as rsync's fuzzy basis. Transfers
 * always write a separate file; the basis remains available for the other
 * direction. Missing rsync and Windows endpoints use compressed scp.
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
      const remote = `${host}:${canUseRsync ? shellQuote(remotePath) : remotePath}`;
      const endpoints = direction === 'download' ? [remote, localPath] : [localPath, remote];
      if (!canUseRsync) {
        deps.runScp(endpoints[0]!, endpoints[1]!);
        return;
      }
      // --checksum prevents a same-size, same-mtime snapshot being skipped.
      // Avoid --inplace: neither a failed transfer nor a matching basis may
      // modify the snapshot we still need to send in the opposite direction.
      const result = deps.runRsync([...RSYNC_OPTIONS, '--quiet', '--', ...endpoints]);
      if (result.error) throw new Error(`Failed to run rsync: ${result.error.message}`);
      if (result.status !== 0) {
        throw new Error(`rsync ${direction} failed for ${host}: ${result.stderr.trim()}`);
      }
    },
  };
}
