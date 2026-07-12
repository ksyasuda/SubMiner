// bun:sqlite bindings for the shared stats-sync engine. The engine itself
// lives in src/core/services/stats-sync so the Electron app (libsql) can run
// the exact same snapshot/merge code via its own driver.
import {
  createDbSnapshot as createDbSnapshotWith,
  createEmptyMergeSummary,
  findLiveStatsDaemonPid,
  nowDbTimestamp,
  readSchemaVersion,
  assertMergeableSchema,
  insertRow,
  tableExists,
  SCHEMA_VERSION,
} from '../../src/core/services/stats-sync/shared.js';
import { openBunSyncDb } from './bun-driver.js';

export {
  createEmptyMergeSummary,
  findLiveStatsDaemonPid,
  nowDbTimestamp,
  readSchemaVersion,
  assertMergeableSchema,
  insertRow,
  tableExists,
  SCHEMA_VERSION,
};
export type { SyncMergeSummary } from '../../src/core/services/stats-sync/shared.js';

export function createDbSnapshot(dbPath: string, outPath: string): void {
  createDbSnapshotWith(openBunSyncDb, dbPath, outPath);
}
