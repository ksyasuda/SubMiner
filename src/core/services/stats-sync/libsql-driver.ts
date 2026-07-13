import { Database } from '../immersion-tracker/sqlite';
import type { SyncDbOpenOptions } from './wal-retry';

export interface SyncDbRunResult {
  changes: number;
  lastInsertRowid: number | bigint;
}

export interface SyncDbStatement {
  run(...params: unknown[]): SyncDbRunResult;
  get(...params: unknown[]): unknown;
  all(...params: unknown[]): unknown[];
}

export interface SyncDb {
  /** Prepare (or reuse a cached prepared statement for) the given SQL. */
  query(sql: string): SyncDbStatement;
  /** Execute SQL that returns no rows (pragmas, transaction control). */
  exec(sql: string): void;
  close(): void;
}

export type SqlRow = Record<string, unknown>;

export function selectAll(db: SyncDb, sql: string, params: unknown[] = []): SqlRow[] {
  return db.query(sql).all(...params) as SqlRow[];
}

export function selectOne(db: SyncDb, sql: string, params: unknown[] = []): SqlRow | undefined {
  return (db.query(sql).get(...params) ?? undefined) as SqlRow | undefined;
}

interface LibsqlDatabase {
  prepare(sql: string): SyncDbStatement;
  exec(sql: string): unknown;
  close(): unknown;
}

/**
 * libsql (better-sqlite3 API) SQLite connection for the stats-sync engine.
 * prepare() is not cached by libsql, so query() keeps a per-connection
 * statement cache — the merge prepares a handful of statements and runs them
 * once per copied row, so re-preparing would dominate merge time.
 */
export function openLibsqlSyncDb(dbPath: string, options: SyncDbOpenOptions): SyncDb {
  const db = new Database(dbPath, {
    readonly: options.readonly === true,
    fileMustExist: options.create !== true,
  }) as unknown as LibsqlDatabase;
  const statements = new Map<string, SyncDbStatement>();
  return {
    query(sql: string): SyncDbStatement {
      let statement = statements.get(sql);
      if (!statement) {
        statement = db.prepare(sql);
        statements.set(sql, statement);
      }
      return statement;
    },
    exec(sql: string): void {
      db.exec(sql);
    },
    close(): void {
      statements.clear();
      db.close();
    },
  };
}
