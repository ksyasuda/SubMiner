// Minimal ambient typing for bun:sqlite. The launcher always runs under bun
// (see the build banner in package.json), but the repo typechecks with plain
// tsc which has no bun type definitions.
declare module 'bun:sqlite' {
  export interface RunResult {
    changes: number;
    lastInsertRowid: number | bigint;
  }

  export interface Statement<ReturnType = unknown, ParamsType extends unknown[] = unknown[]> {
    all(...params: ParamsType): ReturnType[];
    get(...params: ParamsType): ReturnType | undefined;
    run(...params: ParamsType): RunResult;
  }

  export class Database {
    constructor(
      filename: string,
      options?: { readonly?: boolean; readwrite?: boolean; create?: boolean },
    );
    query<ReturnType = unknown, ParamsType extends unknown[] = unknown[]>(
      sql: string,
    ): Statement<ReturnType, ParamsType>;
    prepare<ReturnType = unknown, ParamsType extends unknown[] = unknown[]>(
      sql: string,
    ): Statement<ReturnType, ParamsType>;
    run(sql: string, ...params: unknown[]): RunResult;
    close(throwOnError?: boolean): void;
  }
}
