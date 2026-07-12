import { Database } from '../immersion-tracker/sqlite';
import type { OpenSyncDb, SyncDb, SyncDbRunResult, SyncDbStatement } from './driver';

interface LibsqlStatement {
  run(...params: unknown[]): { changes: number; lastInsertRowid: number | bigint };
  get(...params: unknown[]): unknown;
  all(...params: unknown[]): unknown[];
}

interface LibsqlDatabase {
  prepare(sql: string): LibsqlStatement;
  exec(sql: string): unknown;
  close(): unknown;
}

/**
 * libsql (better-sqlite3 API) binding of the SyncDb driver for the Electron
 * app. prepare() is not cached by libsql, so query() keeps a per-connection
 * statement cache — the merge prepares a handful of statements and runs them
 * once per copied row.
 */
export const openLibsqlSyncDb: OpenSyncDb = (dbPath, options): SyncDb => {
  const db = new Database(dbPath, {
    readonly: options.readonly === true,
    fileMustExist: options.create !== true,
  }) as unknown as LibsqlDatabase;
  const statements = new Map<string, SyncDbStatement>();
  return {
    query(sql: string): SyncDbStatement {
      const cached = statements.get(sql);
      if (cached) return cached;
      const statement = db.prepare(sql);
      const wrapped: SyncDbStatement = {
        run(...params: unknown[]): SyncDbRunResult {
          return statement.run(...params);
        },
        get(...params: unknown[]): unknown {
          return statement.get(...params);
        },
        all(...params: unknown[]): unknown[] {
          return statement.all(...params);
        },
      };
      statements.set(sql, wrapped);
      return wrapped;
    },
    exec(sql: string): void {
      db.exec(sql);
    },
    close(): void {
      statements.clear();
      db.close();
    },
  };
};
