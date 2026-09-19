import type { DatabaseSync } from './sqlite';
import type { TmdbMediaType } from '../../../shared/media-kind';
import { mergeAnimeRecordsInTransaction } from './anime-merge';
import { recomputeLifetimeAnimeAggregatesInTransaction } from './lifetime';
import { toDbTimestamp } from './query-shared';
import { nowMs } from './time';

export interface LiveActionTitleInput {
  tmdbId: number;
  tmdbType: TmdbMediaType;
  titleEnglish: string | null;
  titleNative: string | null;
  description: string | null;
  episodesTotal: number | null;
}

export interface LiveActionLinkResult {
  /** Library entry that carries the TMDB link once the call finishes. */
  animeId: number;
  /** Entries folded into `animeId` because they pointed at the same TMDB title. */
  mergedAnimeIds: number[];
}

export interface LiveActionLinkOptions {
  /**
   * `manual`: the user picked this title, so stored titles are overwritten and
   * every other holder of the TMDB id is folded into this entry.
   * `auto`: an exact filename match, so gaps are filled and the entry joins an
   * existing holder rather than displacing it.
   */
  mode: 'manual' | 'auto';
}

export interface VideoTmdbLink {
  animeId: number;
  tmdbId: number;
  tmdbType: TmdbMediaType;
}

function findOtherTmdbHolders(
  db: DatabaseSync,
  animeId: number,
  input: Pick<LiveActionTitleInput, 'tmdbId' | 'tmdbType'>,
): number[] {
  return (
    db
      .prepare(
        `SELECT anime_id AS animeId
         FROM imm_anime
         WHERE tmdb_id = ? AND tmdb_type = ? AND anime_id != ?
         ORDER BY anime_id ASC`,
      )
      .all(input.tmdbId, input.tmdbType, animeId) as Array<{ animeId: number }>
  ).map((row) => row.animeId);
}

/**
 * Link a library entry to a TMDB title. Unlike AniList, a TMDB show spans all
 * of its seasons, so entries that resolve to the same title are one show and
 * are merged regardless of the season each was parsed with.
 */
export function linkAnimeToTmdbTitle(
  db: DatabaseSync,
  animeId: number,
  input: LiveActionTitleInput,
  options: LiveActionLinkOptions,
): LiveActionLinkResult {
  db.exec('BEGIN IMMEDIATE');
  try {
    const result = linkAnimeToTmdbTitleInTransaction(db, animeId, input, options);
    db.exec('COMMIT');
    return result;
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

/** Caller owns the write transaction, including any artwork replacement. */
export function linkAnimeToTmdbTitleInTransaction(
  db: DatabaseSync,
  animeId: number,
  input: LiveActionTitleInput,
  options: LiveActionLinkOptions,
): LiveActionLinkResult {
  const target = db.prepare('SELECT anilist_id FROM imm_anime WHERE anime_id = ?').get(animeId) as
    | { anilist_id: number | null }
    | undefined;
  if (!target) throw new Error('Unknown library entry');
  if (target.anilist_id !== null) {
    if (options.mode === 'auto')
      throw new Error('Cannot automatically replace an AniList identity');
    // An explicit reassignment changes providers before compatible rows merge.
    db.prepare('UPDATE imm_anime SET anilist_id = NULL WHERE anime_id = ?').run(animeId);
  }
  const others = findOtherTmdbHolders(db, animeId, input);
  let survivor = animeId;
  let mergedAnimeIds: number[] = [];
  if (others.length > 0) {
    if (options.mode === 'manual') {
      mergedAnimeIds = mergeAnimeRecordsInTransaction(db, animeId, others).mergedAnimeIds;
    } else {
      // Keep the entry the user already sees; the newcomer is the transient
      // "Show Season 3" row that a fresh season folder just created.
      survivor = others[0]!;
      mergedAnimeIds = mergeAnimeRecordsInTransaction(db, survivor, [
        animeId,
        ...others.slice(1),
      ]).mergedAnimeIds;
    }
  }

  const updatedAt = toDbTimestamp(nowMs());
  if (options.mode === 'manual') {
    db.prepare(
      `UPDATE imm_anime
       SET media_kind = 'live_action',
           tmdb_id = ?,
           tmdb_type = ?,
           anilist_id = NULL,
           title_romaji = NULL,
           title_english = ?,
           title_native = ?,
           episodes_total = ?,
           description = ?,
           LAST_UPDATE_DATE = ?
       WHERE anime_id = ?`,
    ).run(
      input.tmdbId,
      input.tmdbType,
      input.titleEnglish,
      input.titleNative,
      input.episodesTotal,
      input.description,
      updatedAt,
      survivor,
    );
  } else {
    db.prepare(
      `UPDATE imm_anime
       SET media_kind = 'live_action',
           tmdb_id = ?,
           tmdb_type = ?,
           title_english = COALESCE(title_english, ?),
           title_native = COALESCE(title_native, ?),
           episodes_total = COALESCE(episodes_total, ?),
           description = COALESCE(description, ?),
           LAST_UPDATE_DATE = ?
       WHERE anime_id = ?`,
    ).run(
      input.tmdbId,
      input.tmdbType,
      input.titleEnglish,
      input.titleNative,
      input.episodesTotal,
      input.description,
      updatedAt,
      survivor,
    );
  }
  recomputeLifetimeAnimeAggregatesInTransaction(db);
  return { animeId: survivor, mergedAnimeIds };
}

/** The TMDB link of the live-action entry a video belongs to, if any. */
export function getVideoTmdbLink(db: DatabaseSync, videoId: number): VideoTmdbLink | null {
  const row = db
    .prepare(
      `SELECT a.anime_id AS animeId, a.tmdb_id AS tmdbId, a.tmdb_type AS tmdbType
       FROM imm_videos v
       JOIN imm_anime a ON a.anime_id = v.anime_id
       WHERE v.video_id = ?
         AND a.media_kind = 'live_action'
         AND a.tmdb_id IS NOT NULL
         AND a.tmdb_type IN ('tv', 'movie')`,
    )
    .get(videoId) as VideoTmdbLink | undefined;
  return row ? { animeId: row.animeId, tmdbId: row.tmdbId, tmdbType: row.tmdbType } : null;
}
