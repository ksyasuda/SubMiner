// Minimal ambient typing for bun:sqlite. The launcher always runs under bun
// (see the build banner in package.json), but the repo typechecks with plain
// tsc which has no bun type definitions.
declare module 'bun:sqlite' {
  export interface RunResult {
    changes: number;
    lastInsertRowid: number | bigint;
  }

  export interface Statement {
    all(...params: unknown[]): unknown[];
    get(...params: unknown[]): unknown;
    run(...params: unknown[]): RunResult;
  }

  export class Database {
    constructor(
      filename: string,
      options?: { readonly?: boolean; readwrite?: boolean; create?: boolean },
    );
    query(sql: string): Statement;
    prepare(sql: string): Statement;
    run(sql: string, ...params: unknown[]): RunResult;
    close(): void;
  }
}
