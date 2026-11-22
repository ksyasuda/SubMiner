import { walkFranchiseRelations } from '../anilist/franchise-graph';
import type { AnilistRelationFetcher } from '../anilist/franchise-relation-fetcher';
import { readAnilistRelationCache, writeAnilistRelationNodes } from './anilist-relation-cache';
import type { DatabaseSync } from './sqlite';

/** Cached AniList entries are asked about again after a week. */
export const FRANCHISE_RELATION_STALE_MS = 7 * 24 * 60 * 60 * 1000;
/** Hops followed from a library entry; keeps long-running franchises bounded. */
export const FRANCHISE_WALK_MAX_DEPTH = 6;
/** Entries fetched per refresh (five AniList requests); the rest wait for the next refresh. */
export const FRANCHISE_WALK_MAX_FETCH = 250;
/** Unforced refreshes (library polls) run at most this often. */
const MIN_INTERVAL_MS = 10 * 60 * 1000;

interface RefresherLogger {
  info(message: string, ...args: unknown[]): void;
  warn(message: string, ...args: unknown[]): void;
}

export interface FranchiseRelationRefresherDeps {
  db: DatabaseSync;
  fetcher: AnilistRelationFetcher;
  logger: RefresherLogger;
  now?: () => number;
}

export interface FranchiseRelationRefresher {
  /**
   * Starts a background refresh unless one ran recently. `force` skips that
   * throttle, for when an entry has just been linked to AniList.
   */
  requestRefresh(options?: { force?: boolean }): void;
  /** Resolves once no refresh is running; for tests and shutdown. */
  whenIdle(): Promise<void>;
  /** Stops writing; a running refresh finishes its request and discards the result. */
  stop(): void;
}

function readSeeds(db: DatabaseSync): number[] {
  return (
    db
      .prepare(
        `SELECT anilist_id AS id FROM imm_anime WHERE media_kind = 'anime' AND anilist_id IS NOT NULL
         UNION
         SELECT group_override AS id FROM imm_anime WHERE group_override > 0`,
      )
      .all() as Array<{ id: number }>
  ).map((row) => row.id);
}

export function createFranchiseRelationRefresher(
  deps: FranchiseRelationRefresherDeps,
): FranchiseRelationRefresher {
  const now = deps.now ?? Date.now;
  let running: Promise<void> | null = null;
  let pending = false;
  let stopped = false;
  let lastStartedAtMs: number | null = null;

  const refreshOnce = async (): Promise<void> => {
    const known = readAnilistRelationCache(deps.db);
    const result = await walkFranchiseRelations({
      seeds: readSeeds(deps.db),
      known,
      isFresh: (node) => now() - node.fetchedAtMs < FRANCHISE_RELATION_STALE_MS,
      fetchNodes: (ids) => deps.fetcher.fetchMedia(ids),
      onBatch: (nodes) => writeAnilistRelationNodes(deps.db, nodes),
      isCancelled: () => stopped,
      now,
      maxDepth: FRANCHISE_WALK_MAX_DEPTH,
      maxFetch: FRANCHISE_WALK_MAX_FETCH,
    });
    if (result.fetched > 0) {
      deps.logger.info(
        'franchise-relations: refreshed %d AniList entries%s',
        result.fetched,
        result.truncated ? ' (more next time)' : '',
      );
    }
  };

  const run = (): void => {
    lastStartedAtMs = now();
    running = (async () => {
      try {
        do {
          pending = false;
          await refreshOnce();
        } while (pending && !stopped);
      } catch (error) {
        // Cached rows stay as they were; the next request (after the throttle) retries.
        deps.logger.warn(
          'franchise-relations: refresh failed: %s',
          error instanceof Error ? error.message : String(error),
        );
      } finally {
        running = null;
      }
    })();
  };

  return {
    requestRefresh(options = {}) {
      if (stopped) return;
      if (running) {
        if (options.force) pending = true;
        return;
      }
      if (!options.force && lastStartedAtMs !== null && now() - lastStartedAtMs < MIN_INTERVAL_MS) {
        return;
      }
      run();
    },
    async whenIdle() {
      while (running) await running;
    },
    stop() {
      stopped = true;
    },
  };
}
