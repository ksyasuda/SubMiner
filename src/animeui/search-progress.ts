import type { AnimeBrowserEntry, AnimeBrowserSearchUpdate } from '../types/anime-browser';
import type { SourceSearchFailure } from '../types/anime-browser';

/**
 * Streamed-search bookkeeping, kept apart from the DOM so the staleness rules
 * are testable.
 *
 * Updates stream in while earlier searches may still be resolving; a search is
 * identified by its token and only the newest one may touch the grid. Tokens
 * are emitted in start order over an ordered channel, so "newest" is simply
 * the highest token seen.
 */

export interface SearchProgress {
  token: number;
  sourceCount: number;
  sourcesDone: number;
  entryCount: number;
  failures: SourceSearchFailure[];
  done: boolean;
}

const IDLE: SearchProgress = {
  token: -1,
  sourceCount: 0,
  sourcesDone: 0,
  entryCount: 0,
  failures: [],
  done: false,
};

export interface AppliedUpdate {
  progress: SearchProgress;
  /** Entries this update contributed; append them to the grid. */
  entries: AnimeBrowserEntry[];
  /** True when this update began a new search; clear the grid first. */
  started: boolean;
}

export function idleSearchProgress(): SearchProgress {
  return { ...IDLE, failures: [] };
}

/**
 * Fold one update into the current progress. Returns null for an update from
 * a superseded search, which the caller must ignore entirely.
 */
export function applySearchUpdate(
  current: SearchProgress,
  update: AnimeBrowserSearchUpdate,
): AppliedUpdate | null {
  if (update.kind === 'start') {
    // An older search's start (or a replay) must not reset the newer one.
    if (update.token <= current.token) return null;
    return {
      progress: {
        token: update.token,
        sourceCount: update.sourceCount,
        sourcesDone: 0,
        entryCount: 0,
        failures: [],
        done: false,
      },
      entries: [],
      started: true,
    };
  }

  if (update.token !== current.token) return null;

  if (update.kind === 'result') {
    return {
      progress: {
        ...current,
        sourcesDone: current.sourcesDone + 1,
        entryCount: current.entryCount + update.entries.length,
      },
      entries: update.entries,
      started: false,
    };
  }

  if (update.kind === 'failure') {
    return {
      progress: {
        ...current,
        sourcesDone: current.sourcesDone + 1,
        failures: [...current.failures, update.failure],
      },
      entries: [],
      started: false,
    };
  }

  return { progress: { ...current, done: true }, entries: [], started: false };
}

/** `Searching… 3/5 sources · 42 results`, with failures named once they exist. */
export function summarizeProgress(progress: SearchProgress): string {
  const counts = `${progress.sourcesDone}/${progress.sourceCount} sources · ${progress.entryCount} result${progress.entryCount === 1 ? '' : 's'}`;
  const failed =
    progress.failures.length === 0
      ? ''
      : ` · unavailable: ${progress.failures.map((failure) => failure.sourceName).join(', ')}`;
  return `Searching… ${counts}${failed}`;
}
