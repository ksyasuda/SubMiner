import type { DatabaseSync } from './sqlite';
import {
  animeSeasonsAreMergeCompatible,
  getParsedSeasonsForAnime,
  mergeAnimeRecordsInTransaction,
} from './anime-merge';
import {
  hasExactStoredTitleMatch,
  hasDismissedAnimeMergeRecommendation,
  recordAnimeMergeRecommendation,
  shouldRecommendAnilistConflict,
  type AnimeConflictRecommendationOptions,
} from './anime-merge-recommendations';
import { getOrCreateAnimeRecord } from './storage';
import { toDbTimestamp } from './query-shared';
import { nowMs } from './time';

export interface AnimeSeasonRepairSummary {
  scanned: number;
  repaired: number;
  movedVideos: number;
  deletedAnimeRows: number;
  /**
   * Entry that owns the videos afterwards when two rows were folded together,
   * so callers can keep pointing at a row that still exists.
   */
  survivingAnimeId: number | null;
  /** True when an ambiguous AniList collision was saved for user review. */
  mergeRecommended: boolean;
  /** True when automatic metadata must not assign the colliding AniList id. */
  anilistAssignmentBlocked: boolean;
}

export interface AnimeAnilistConflictOptions extends AnimeConflictRecommendationOptions {
  /**
   * Which row keeps its identity when two entries claim the same AniList id.
   * `existing` (the default) keeps the row that already held the id, so
   * automatic cover-art resolution does not rename a card under the user;
   * `target` keeps the row the user is acting on.
   */
  survivor?: 'target' | 'existing';
}

export {
  dismissAnimeMergeRecommendation,
  getAnimeMergeRecommendations,
  type AnimeMergeRecommendation,
} from './anime-merge-recommendations';

interface AnimeRow {
  anime_id: number;
  anilist_id: number | null;
  title_romaji: string | null;
  title_english: string | null;
  title_native: string | null;
  episodes_total: number | null;
  description: string | null;
}

interface ParsedVideoRow {
  video_id: number;
  parsed_title: string | null;
  parsed_season: number | null;
  anime_assignment_locked: number;
}

interface RedistributeOptions {
  transferAnilistToAnimeId?: number | null;
  transferLegacyAnilist?: boolean;
  overwriteTargetAnilist?: boolean;
}

function emptySummary(scanned = 0): AnimeSeasonRepairSummary {
  return {
    scanned,
    repaired: 0,
    movedVideos: 0,
    deletedAnimeRows: 0,
    survivingAnimeId: null,
    mergeRecommended: false,
    anilistAssignmentBlocked: false,
  };
}

function mergeSummary(
  target: AnimeSeasonRepairSummary,
  source: AnimeSeasonRepairSummary,
): AnimeSeasonRepairSummary {
  target.scanned += source.scanned;
  target.repaired += source.repaired;
  target.movedVideos += source.movedVideos;
  target.deletedAnimeRows += source.deletedAnimeRows;
  target.survivingAnimeId = source.survivingAnimeId ?? target.survivingAnimeId;
  target.mergeRecommended ||= source.mergeRecommended;
  target.anilistAssignmentBlocked ||= source.anilistAssignmentBlocked;
  return target;
}

function runInTransaction<T>(db: DatabaseSync, work: () => T): T {
  db.exec('BEGIN');
  try {
    const result = work();
    db.exec('COMMIT');
    return result;
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

function normalizeSeason(value: number | null): number | null {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value <= 0) {
    return null;
  }
  return value;
}

function getAnimeRow(db: DatabaseSync, animeId: number): AnimeRow | null {
  return db
    .prepare(
      `
        SELECT
          anime_id,
          anilist_id,
          title_romaji,
          title_english,
          title_native,
          episodes_total,
          description
        FROM imm_anime
        WHERE anime_id = ?
      `,
    )
    .get(animeId) as AnimeRow | null;
}

function getParsedVideos(db: DatabaseSync, animeId: number): ParsedVideoRow[] {
  return db
    .prepare(
      `
        SELECT video_id, parsed_title, parsed_season, anime_assignment_locked
        FROM imm_videos
        WHERE anime_id = ?
        ORDER BY video_id ASC
      `,
    )
    .all(animeId) as ParsedVideoRow[];
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

function assignAnilistToTarget(
  db: DatabaseSync,
  source: AnimeRow,
  targetAnimeId: number,
  overwriteTarget: boolean,
  updatedAt: string,
): boolean {
  if (source.anilist_id === null || targetAnimeId === source.anime_id) {
    return false;
  }

  const target = getAnimeRow(db, targetAnimeId);
  if (!target) {
    return false;
  }
  if (!overwriteTarget && target.anilist_id !== null && target.anilist_id !== source.anilist_id) {
    return false;
  }

  db.prepare(
    `
      UPDATE imm_anime
      SET anilist_id = NULL,
          LAST_UPDATE_DATE = ?
      WHERE anime_id = ?
    `,
  ).run(updatedAt, source.anime_id);

  const updated = db
    .prepare(
      `
        UPDATE imm_anime
        SET
          anilist_id = ?,
          title_romaji = COALESCE(?, title_romaji),
          title_english = COALESCE(?, title_english),
          title_native = COALESCE(?, title_native),
          episodes_total = COALESCE(?, episodes_total),
          description = COALESCE(?, description),
          LAST_UPDATE_DATE = ?
        WHERE anime_id = ?
      `,
    )
    .run(
      source.anilist_id,
      source.title_romaji,
      source.title_english,
      source.title_native,
      source.episodes_total,
      source.description,
      updatedAt,
      targetAnimeId,
    ) as { changes: number };
  return updated.changes > 0;
}

function redistributeAnimeRowByParsedSeasonsInTransaction(
  db: DatabaseSync,
  animeId: number,
  options: RedistributeOptions = {},
): AnimeSeasonRepairSummary {
  const source = getAnimeRow(db, animeId);
  if (!source) {
    return emptySummary(1);
  }

  const videos = getParsedVideos(db, animeId);
  const summary = emptySummary(1);
  const updatedAt = toDbTimestamp(nowMs());
  const targetBySeason = new Map<number, number>();

  for (const video of videos) {
    if (video.anime_assignment_locked === 1) {
      continue;
    }
    const parsedTitle = video.parsed_title?.trim();
    const season = normalizeSeason(video.parsed_season);
    if (!parsedTitle || season === null) {
      continue;
    }

    const targetAnimeId = getOrCreateAnimeRecord(db, {
      parsedTitle,
      canonicalTitle: parsedTitle,
      seasonScope: season,
      anilistId: null,
      titleRomaji: null,
      titleEnglish: null,
      titleNative: null,
      metadataJson: null,
    });
    targetBySeason.set(season, targetAnimeId);

    if (targetAnimeId === animeId) {
      continue;
    }

    const videoUpdate = db
      .prepare(
        `
          UPDATE imm_videos
          SET anime_id = ?,
              LAST_UPDATE_DATE = ?
          WHERE video_id = ?
        `,
      )
      .run(targetAnimeId, updatedAt, video.video_id) as { changes: number };
    const lineUpdate = db
      .prepare(
        `
          UPDATE imm_subtitle_lines
          SET anime_id = ?,
              LAST_UPDATE_DATE = ?
          WHERE video_id = ?
        `,
      )
      .run(targetAnimeId, updatedAt, video.video_id) as { changes: number };

    if (videoUpdate.changes > 0 || lineUpdate.changes > 0) {
      summary.movedVideos += 1;
    }
  }

  const transferTarget =
    options.transferAnilistToAnimeId ??
    (options.transferLegacyAnilist
      ? (targetBySeason.get(1) ??
        (targetBySeason.size === 1 ? [...targetBySeason.values()][0] : null))
      : null);
  if (transferTarget) {
    const transferred = assignAnilistToTarget(
      db,
      source,
      transferTarget,
      options.overwriteTargetAnilist ?? false,
      updatedAt,
    );
    if (transferred) {
      summary.repaired += 1;
    }
  }

  if (!hasAnimeReferences(db, animeId)) {
    const deleted = db.prepare('DELETE FROM imm_anime WHERE anime_id = ?').run(animeId) as {
      changes: number;
    };
    if (deleted.changes > 0) {
      summary.deletedAnimeRows += 1;
    }
  }

  if (summary.movedVideos > 0 || summary.deletedAnimeRows > 0) {
    summary.repaired += 1;
  }
  return summary;
}

export function repairLegacySeasonlessAnimeRows(db: DatabaseSync): AnimeSeasonRepairSummary {
  return runInTransaction(db, () => {
    const candidates = db
      .prepare(
        `
          SELECT a.anime_id AS animeId
          FROM imm_anime a
          JOIN imm_videos v ON v.anime_id = a.anime_id
          WHERE v.parsed_title IS NOT NULL
            AND TRIM(v.parsed_title) != ''
            AND v.parsed_season IS NOT NULL
            AND v.parsed_season > 0
          GROUP BY a.anime_id
          HAVING COUNT(DISTINCT v.parsed_season) > 1
          ORDER BY a.anime_id ASC
        `,
      )
      .all() as Array<{ animeId: number }>;
    const summary = emptySummary();
    for (const candidate of candidates) {
      mergeSummary(
        summary,
        redistributeAnimeRowByParsedSeasonsInTransaction(db, candidate.animeId, {
          transferLegacyAnilist: true,
        }),
      );
    }
    return summary;
  });
}

/**
 * Two library entries cannot both hold the same AniList id
 * (`imm_anime.anilist_id` is UNIQUE). Fold an automatic collision only when
 * exact title evidence and compatible parsed seasons make it safe. Persist a
 * review recommendation for compatible weak matches. Legacy season
 * redistribution of a conflicting multi-season row is reserved for exact or
 * manual matches; a weak automatic match blocks the assignment instead.
 */
export function resolveAnimeAnilistConflict(
  db: DatabaseSync,
  targetAnimeId: number,
  anilistId: number,
  options: AnimeAnilistConflictOptions = {},
): AnimeSeasonRepairSummary {
  const conflict = db
    .prepare(
      `
        SELECT anime_id AS animeId
        FROM imm_anime
        WHERE anilist_id = ?
          AND anime_id != ?
        LIMIT 1
      `,
    )
    .get(anilistId, targetAnimeId) as { animeId: number } | null;
  if (!conflict) {
    return emptySummary();
  }

  return runInTransaction(db, () => {
    const targetRow = getAnimeRow(db, targetAnimeId);
    if (
      options.survivor !== 'target' &&
      targetRow?.anilist_id != null &&
      targetRow.anilist_id !== anilistId
    ) {
      // An automatic lookup disagreeing with an existing explicit link is a
      // mis-resolution, not evidence that either row should move or merge. The
      // colliding id must not be assigned either: another row owns it and
      // imm_anime.anilist_id is UNIQUE.
      const summary = emptySummary(1);
      summary.anilistAssignmentBlocked = true;
      return summary;
    }
    const isManual = options.survivor === 'target' || options.matchConfidence === 'manual';
    if (!isManual && hasDismissedAnimeMergeRecommendation(db, targetAnimeId, conflict.animeId)) {
      const summary = emptySummary(1);
      summary.anilistAssignmentBlocked = true;
      return summary;
    }
    const targetSeasons = getParsedSeasonsForAnime(db, targetAnimeId);
    const conflictSeasons = getParsedSeasonsForAnime(db, conflict.animeId);
    if (
      !isManual &&
      targetSeasons.size === 1 &&
      conflictSeasons.size === 1 &&
      [...targetSeasons][0] !== [...conflictSeasons][0]
    ) {
      const summary = emptySummary(1);
      summary.anilistAssignmentBlocked = true;
      return summary;
    }
    if (canMergeAnilistConflict(db, targetAnimeId, conflict.animeId, anilistId, options)) {
      const survivingAnimeId = options.survivor === 'target' ? targetAnimeId : conflict.animeId;
      const absorbedAnimeId = survivingAnimeId === targetAnimeId ? conflict.animeId : targetAnimeId;
      const merge = mergeAnimeRecordsInTransaction(db, survivingAnimeId, [absorbedAnimeId]);
      const summary = emptySummary(1);
      summary.movedVideos = merge.movedVideos;
      summary.deletedAnimeRows = merge.mergedAnimeIds.length;
      if (merge.mergedAnimeIds.length > 0) {
        summary.repaired = 1;
        // Only reported once a row really absorbed the other, so callers never
        // follow this to an anime id that was never written.
        summary.survivingAnimeId = survivingAnimeId;
      }
      // Lifetime summaries are rebuilt by the caller off this summary, the same
      // as the redistribution path below.
      return summary;
    }

    if (shouldRecommendAnilistConflict(db, targetAnimeId, conflict.animeId, options)) {
      recordAnimeMergeRecommendation(db, targetAnimeId, conflict.animeId, anilistId);
      const summary = emptySummary(1);
      summary.mergeRecommended = true;
      return summary;
    }

    const isExactAutomaticMatch =
      options.matchConfidence === 'exact' ||
      (options.matchConfidence === undefined &&
        hasExactStoredTitleMatch(db, targetAnimeId, conflict.animeId));
    if (!isManual && !isExactAutomaticMatch) {
      // Redistribution dismantles the id's current owner and hands the id to
      // the target. On a weak automatic match that owner is usually the
      // correctly linked card (e.g. a legitimate multi-season entry), so
      // splitting it here is exactly the fuzzy false merge this gate exists to
      // stop. Only exact or manual evidence may fall through.
      const summary = emptySummary(1);
      summary.anilistAssignmentBlocked = true;
      return summary;
    }

    return redistributeAnimeRowByParsedSeasonsInTransaction(db, conflict.animeId, {
      transferAnilistToAnimeId: targetAnimeId,
      overwriteTargetAnilist: true,
    });
  });
}

function canMergeAnilistConflict(
  db: DatabaseSync,
  targetAnimeId: number,
  conflictAnimeId: number,
  anilistId: number,
  options: AnimeAnilistConflictOptions,
): boolean {
  const targetRow = getAnimeRow(db, targetAnimeId);
  if (!targetRow) {
    // Nothing to merge with a row that no longer exists (a stale id from the
    // caller); fall through to the redistribution path.
    return false;
  }
  if (options.survivor !== 'target') {
    // The target is the row about to disappear here, so an existing link of its
    // own means this is a mis-resolution rather than a duplicate: leave it be.
    if (targetRow.anilist_id != null && targetRow.anilist_id !== anilistId) {
      return false;
    }
  }
  if (
    options.matchConfidence === 'weak' ||
    (options.matchConfidence === undefined &&
      !hasExactStoredTitleMatch(db, targetAnimeId, conflictAnimeId))
  ) {
    return false;
  }
  return animeSeasonsAreMergeCompatible(
    getParsedSeasonsForAnime(db, targetAnimeId),
    getParsedSeasonsForAnime(db, conflictAnimeId),
  );
}
