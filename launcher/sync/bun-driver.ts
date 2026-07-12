import { Database } from 'bun:sqlite';
import type {
  OpenSyncDb,
  SyncDb,
  SyncDbStatement,
} from '../../src/core/services/stats-sync/driver.js';

// bun:sqlite's Database.query() already caches prepared statements per SQL
// string, which is exactly what the SyncDb contract asks for.
export const openBunSyncDb: OpenSyncDb = (dbPath, options): SyncDb => {
  const db = new Database(dbPath, options);
  return {
    query(sql: string): SyncDbStatement {
      return db.query(sql);
    },
    exec(sql: string): void {
      db.run(sql);
    },
    close(): void {
      db.close();
    },
  };
};
