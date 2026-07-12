// Minimal SQLite driver surface the stats-sync engine runs against. The
// launcher binds it to bun:sqlite; the Electron app binds it to libsql. Both
// implementations must cache prepared statements per SQL string in query():
// the merge runs one insert per copied row, so re-preparing would dominate
// merge time.

export interface SyncDbRunResult {
  changes: number;
  lastInsertRowid: number | bigint;
}

export interface SyncDbStatement {
  run(...params: unknown[]): SyncDbRunResult;
  get(...params: unknown[]): unknown;
  all(...params: unknown[]): unknown[];
}

export interface SyncDbOpenOptions {
  readonly?: boolean;
  readwrite?: boolean;
  create?: boolean;
}

export interface SyncDb {
  /** Prepare (or reuse a cached prepared statement for) the given SQL. */
  query(sql: string): SyncDbStatement;
  /** Execute SQL that returns no rows (pragmas, transaction control). */
  exec(sql: string): void;
  close(): void;
}

export type OpenSyncDb = (dbPath: string, options: SyncDbOpenOptions) => SyncDb;

export type SqlRow = Record<string, unknown>;

export function selectAll(db: SyncDb, sql: string, params: unknown[] = []): SqlRow[] {
  return db.query(sql).all(...params) as SqlRow[];
}

export function selectOne(db: SyncDb, sql: string, params: unknown[] = []): SqlRow | undefined {
  return (db.query(sql).get(...params) ?? undefined) as SqlRow | undefined;
}
