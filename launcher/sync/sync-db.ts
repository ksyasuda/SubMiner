import { mergeSnapshotIntoDb as mergeSnapshotIntoDbWith } from '../../src/core/services/stats-sync/merge.js';
import { openBunSyncDb } from './bun-driver.js';
import type { SyncMergeSummary } from './sync-shared.js';

export type { SyncMergeSummary } from './sync-shared.js';
export { createDbSnapshot, findLiveStatsDaemonPid } from './sync-shared.js';
export { formatMergeSummary } from '../../src/core/services/stats-sync/merge.js';

/** bun:sqlite binding of the shared snapshot-merge engine. */
export function mergeSnapshotIntoDb(localDbPath: string, snapshotPath: string): SyncMergeSummary {
  return mergeSnapshotIntoDbWith(openBunSyncDb, localDbPath, snapshotPath);
}
