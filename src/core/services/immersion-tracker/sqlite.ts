import Database = require('libsql');

export { Database };

export interface DatabaseRunResult {
  changes: number;
  lastInsertRowid: number | bigint;
}

export interface DatabaseStatement {
  run(...params: unknown[]): DatabaseRunResult;
  get(...params: unknown[]): unknown;
  all(...params: unknown[]): unknown[];
}

export interface DatabaseSync {
  prepare(source: string): DatabaseStatement;
  exec(source: string): DatabaseSync;
  close(): DatabaseSync;
}
