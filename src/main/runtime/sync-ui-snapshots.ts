import fs from 'node:fs';
import path from 'node:path';
import type { SyncUiSnapshotFile, SyncUiStartResult } from '../../types/sync-ui';

interface SyncUiSnapshotsDeps {
  snapshotsDir: string;
  nowMs: () => number;
  startRun: (kind: 'snapshot' | 'merge', host: null, args: string[]) => SyncUiStartResult;
  broadcastStateChanged: () => void;
}

function formatSnapshotName(nowMs: number): string {
  // Milliseconds prevent snapshots taken in the same second from overwriting each other.
  const iso = new Date(nowMs).toISOString();
  const stamp = iso.slice(0, 23).replace(/[-:.]/g, '').replace('T', '-');
  return `immersion-${stamp}.sqlite`;
}

export function createSyncUiSnapshots(deps: SyncUiSnapshotsDeps) {
  function listSnapshots(): SyncUiSnapshotFile[] {
    let names: string[];
    try {
      names = fs.readdirSync(deps.snapshotsDir);
    } catch {
      return [];
    }
    const files: SyncUiSnapshotFile[] = [];
    for (const name of names) {
      if (!name.endsWith('.sqlite')) continue;
      const filePath = path.join(deps.snapshotsDir, name);
      try {
        const stat = fs.statSync(filePath);
        if (!stat.isFile()) continue;
        files.push({
          path: filePath,
          name,
          sizeBytes: stat.size,
          modifiedAtMs: stat.mtimeMs,
        });
      } catch {
        // file disappeared between readdir and stat
      }
    }
    files.sort((a, b) => b.modifiedAtMs - a.modifiedAtMs);
    return files;
  }

  function createSnapshot(): SyncUiStartResult {
    const outPath = path.join(deps.snapshotsDir, formatSnapshotName(deps.nowMs()));
    return deps.startRun('snapshot', null, ['sync', '--snapshot', outPath, '--json']);
  }

  function mergeSnapshotFile(filePath: string, force = false): SyncUiStartResult {
    if (!fs.existsSync(filePath)) {
      return { started: false, runId: null, reason: `Snapshot file not found: ${filePath}` };
    }
    const args = ['sync', '--merge', filePath];
    if (force) args.push('--force');
    args.push('--json');
    return deps.startRun('merge', null, args);
  }

  function deleteSnapshot(filePath: string): void {
    const resolved = path.resolve(filePath);
    const dir = path.resolve(deps.snapshotsDir);
    if (resolved !== path.join(dir, path.basename(resolved)) || !resolved.endsWith('.sqlite')) {
      throw new Error('Refusing to delete a file outside the snapshots directory.');
    }
    fs.rmSync(resolved, { force: true });
    deps.broadcastStateChanged();
  }

  return { listSnapshots, createSnapshot, mergeSnapshotFile, deleteSnapshot };
}
