import type { SourceSearchFailure } from '../types/anime-browser';

/**
 * Running one query against every installed source at once.
 *
 * Each source is a separate extension behind the same single-threaded bridge,
 * so the fan-out is bounded rather than unleashed: a dozen extensions all
 * uploading and searching at once starves the ones the user is waiting on.
 */

/** Enough to hide the latency of a slow source without queueing the bridge. */
const DEFAULT_CONCURRENCY = 4;

export interface SourceTarget {
  id: string;
  name: string;
}

export interface FanOutResult<T> {
  /** One entry per source that succeeded, in source order. */
  results: T[];
  /** One entry per source that threw, in source order. */
  failures: SourceSearchFailure[];
}

/**
 * Run `task` against every source, at most `concurrency` at a time.
 *
 * A source that throws becomes a failure instead of rejecting the whole call —
 * one misconfigured extension must not hide every other source's results.
 */
export async function mapSourcesConcurrently<S extends SourceTarget, T>(
  sources: S[],
  task: (source: S) => Promise<T>,
  concurrency: number = DEFAULT_CONCURRENCY,
): Promise<FanOutResult<T>> {
  // Slots keep the output in source order regardless of completion order, so
  // the same query lays out the same way twice.
  const results: Array<{ value: T } | null> = sources.map(() => null);
  const failures: Array<SourceSearchFailure | null> = sources.map(() => null);
  let next = 0;

  const worker = async (): Promise<void> => {
    for (;;) {
      const index = next;
      next += 1;
      const source = sources[index];
      if (!source) return;
      try {
        results[index] = { value: await task(source) };
      } catch (error) {
        failures[index] = {
          sourceId: source.id,
          sourceName: source.name,
          error: error instanceof Error ? error.message : String(error),
        };
      }
    }
  };

  const workers = Math.max(1, Math.min(concurrency, sources.length));
  await Promise.all(Array.from({ length: workers }, () => worker()));

  return {
    results: results
      .filter((slot): slot is { value: T } => slot !== null)
      .map((slot) => slot.value),
    failures: failures.filter((slot): slot is SourceSearchFailure => slot !== null),
  };
}

/**
 * Round-robin merge, so the grid opens with one hit from each source rather
 * than the whole of the first source before the second one starts.
 */
export function interleave<T>(groups: T[][]): T[] {
  const merged: T[] = [];
  const longest = groups.reduce((max, group) => Math.max(max, group.length), 0);
  for (let index = 0; index < longest; index += 1) {
    for (const group of groups) {
      if (index < group.length) merged.push(group[index] as T);
    }
  }
  return merged;
}
