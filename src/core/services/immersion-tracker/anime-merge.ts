import type { DatabaseSync } from './sqlite';
import { rebuildLifetimeSummariesInTransaction } from './lifetime';
import { toDbTimestamp } from './query-shared';
import { nowMs } from './time';

export interface AnimeMergeSummary {
  /** Library entry that owns every moved episode once the merge finishes. */
  survivingAnimeId: number;
  /** Entries that were folded into the survivor and deleted. */
  mergedAnimeIds: number[];
  movedVideos: number;
}

export interface VideoMoveSummary {
  targetAnimeId: number;
  /** Previous owner, or null when the episode had no library entry yet. */
  previousAnimeId: number | null;
  /** True when the previous owner was left empty and pruned. */
  removedPreviousAnime: boolean;
}

interface AnimeMetadataRow {
  anilist_id: number | null;
  title_romaji: string | null;
  title_english: string | null;
  title_native: string | null;
  episodes_total: number | null;
  description: string | null;
}

function emptyMergeSummary(survivingAnimeId: number): AnimeMergeSummary {
  return { survivingAnimeId, mergedAnimeIds: [], movedVideos: 0 };
}

function runInTransaction<T>(db: DatabaseSync, work: () => T): T {
  db.exec('BEGIN IMMEDIATE');
  try {
    const result = work();
    db.exec('COMMIT');
    return result;
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

function readAnimeMetadata(db: DatabaseSync, animeId: number): AnimeMetadataRow | null {
  return (db
    .prepare(
      `
        SELECT anilist_id, title_romaji, title_english, title_native, episodes_total, description
        FROM imm_anime
        WHERE anime_id = ?
      `,
    )
    .get(animeId) ?? null) as AnimeMetadataRow | null;
}

function animeExists(db: DatabaseSync, animeId: number): boolean {
  return Boolean(db.prepare('SELECT 1 FROM imm_anime WHERE anime_id = ?').get(animeId));
}

function hasAnimeReferences(db: DatabaseSync, animeId: number): boolean {
  const row = db
    .prepare(
      `
        SELECT 1 AS found
        WHERE EXISTS (SELECT 1 FROM imm_videos WHERE anime_id = ?)
           OR EXISTS (SELECT 1 FROM imm_subtitle_lines WHERE anime_id = ?)
      `,
    )
    .get(animeId, animeId) as { found: number } | null;
  return Boolean(row);
}

/**
 * Distinct explicit seasons behind a library entry. Videos with no parsed
 * season are ignored, so an entry built from `Show - 03.mkv` style filenames
 * reports an empty set rather than a bogus season.
 */
export function getParsedSeasonsForAnime(db: DatabaseSync, animeId: number): Set<number> {
  const rows = db
    .prepare(
      `
        SELECT DISTINCT parsed_season AS season
        FROM imm_videos
        WHERE anime_id = ?
          AND parsed_season IS NOT NULL
          AND parsed_season > 0
      `,
    )
    .all(animeId) as Array<{ season: number }>;
  return new Set(rows.map((row) => row.season));
}

/**
 * Two entries are safe to fold together when neither spans more than one
 * explicit season and they do not disagree about which season that is. A
 * seasonless entry is compatible with anything single-season: those are the
 * `Show - 03.mkv` vs `Show.S01E03.mkv` splits that produce duplicate cards.
 */
export function animeSeasonsAreMergeCompatible(a: Set<number>, b: Set<number>): boolean {
  if (a.size > 1 || b.size > 1) return false;
  if (a.size === 0 || b.size === 0) return true;
  return [...a][0] === [...b][0];
}

/**
 * Fill in whatever the target is missing from a source row that is on its way
 * out. Must run after the source row is deleted: imm_anime.anilist_id is
 * UNIQUE, so the two rows cannot hold the same id at once.
 */
function absorbAnimeMetadata(
  db: DatabaseSync,
  targetAnimeId: number,
  source: AnimeMetadataRow | null,
  updatedAt: string,
): void {
  if (!source) return;
  db.prepare(
    `
      UPDATE imm_anime
      SET
        anilist_id = COALESCE(anilist_id, ?),
        title_romaji = COALESCE(title_romaji, ?),
        title_english = COALESCE(title_english, ?),
        title_native = COALESCE(title_native, ?),
        episodes_total = COALESCE(episodes_total, ?),
        description = COALESCE(description, ?),
        LAST_UPDATE_DATE = ?
      WHERE anime_id = ?
    `,
  ).run(
    source.anilist_id,
    source.title_romaji,
    source.title_english,
    source.title_native,
    source.episodes_total,
    source.description,
    updatedAt,
    targetAnimeId,
  );
}

/**
 * Fold `sourceAnimeIds` into `targetAnimeId`: every episode and subtitle line
 * is repointed, metadata the target is missing is inherited from the sources,
 * and the emptied source rows are deleted.
 *
 * Assumes the caller already holds a write transaction and rebuilds the
 * lifetime summaries afterwards; use {@link mergeAnimeRecords} otherwise.
 */
export function mergeAnimeRecordsInTransaction(
  db: DatabaseSync,
  targetAnimeId: number,
  sourceAnimeIds: number[],
): AnimeMergeSummary {
  const summary = emptyMergeSummary(targetAnimeId);
  if (!animeExists(db, targetAnimeId)) {
    return summary;
  }

  const updatedAt = toDbTimestamp(nowMs());
  const moveVideosStmt = db.prepare(
    'UPDATE imm_videos SET anime_id = ?, LAST_UPDATE_DATE = ? WHERE anime_id = ?',
  );
  const moveLinesStmt = db.prepare(
    'UPDATE imm_subtitle_lines SET anime_id = ?, LAST_UPDATE_DATE = ? WHERE anime_id = ?',
  );
  const dropLifetimeStmt = db.prepare('DELETE FROM imm_lifetime_anime WHERE anime_id = ?');
  const dropAnimeStmt = db.prepare('DELETE FROM imm_anime WHERE anime_id = ?');

  for (const sourceAnimeId of new Set(sourceAnimeIds)) {
    if (sourceAnimeId === targetAnimeId || !animeExists(db, sourceAnimeId)) {
      continue;
    }

    const sourceMetadata = readAnimeMetadata(db, sourceAnimeId);
    const moved = moveVideosStmt.run(targetAnimeId, updatedAt, sourceAnimeId) as {
      changes: number;
    };
    moveLinesStmt.run(targetAnimeId, updatedAt, sourceAnimeId);
    dropLifetimeStmt.run(sourceAnimeId);
    dropAnimeStmt.run(sourceAnimeId);
    absorbAnimeMetadata(db, targetAnimeId, sourceMetadata, updatedAt);

    summary.mergedAnimeIds.push(sourceAnimeId);
    summary.movedVideos += moved.changes;
  }

  return summary;
}

export function mergeAnimeRecords(
  db: DatabaseSync,
  targetAnimeId: number,
  sourceAnimeIds: number[],
): AnimeMergeSummary {
  return runInTransaction(db, () => {
    const summary = mergeAnimeRecordsInTransaction(db, targetAnimeId, sourceAnimeIds);
    if (summary.mergedAnimeIds.length > 0) {
      rebuildLifetimeSummariesInTransaction(db);
    }
    return summary;
  });
}

/**
 * Move a single episode to another library entry, pruning the previous owner
 * when it is left with nothing.
 */
export function moveVideoToAnime(
  db: DatabaseSync,
  videoId: number,
  targetAnimeId: number,
): VideoMoveSummary {
  return runInTransaction(db, () => {
    const videoRow = db
      .prepare('SELECT anime_id AS animeId FROM imm_videos WHERE video_id = ?')
      .get(videoId) as { animeId: number | null } | null;
    if (!videoRow || !animeExists(db, targetAnimeId)) {
      throw new Error('Unknown episode or target library entry');
    }

    const previousAnimeId = videoRow.animeId;
    if (previousAnimeId === targetAnimeId) {
      return { targetAnimeId, previousAnimeId, removedPreviousAnime: false };
    }

    const updatedAt = toDbTimestamp(nowMs());
    db.prepare('UPDATE imm_videos SET anime_id = ?, LAST_UPDATE_DATE = ? WHERE video_id = ?').run(
      targetAnimeId,
      updatedAt,
      videoId,
    );
    db.prepare(
      'UPDATE imm_subtitle_lines SET anime_id = ?, LAST_UPDATE_DATE = ? WHERE video_id = ?',
    ).run(targetAnimeId, updatedAt, videoId);

    let removedPreviousAnime = false;
    if (previousAnimeId !== null && !hasAnimeReferences(db, previousAnimeId)) {
      const sourceMetadata = readAnimeMetadata(db, previousAnimeId);
      db.prepare('DELETE FROM imm_lifetime_anime WHERE anime_id = ?').run(previousAnimeId);
      db.prepare('DELETE FROM imm_anime WHERE anime_id = ?').run(previousAnimeId);
      absorbAnimeMetadata(db, targetAnimeId, sourceMetadata, updatedAt);
      removedPreviousAnime = true;
    }

    rebuildLifetimeSummariesInTransaction(db);
    return { targetAnimeId, previousAnimeId, removedPreviousAnime };
  });
}
