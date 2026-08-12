import type { DatabaseSync } from './sqlite';
import { animeSeasonsAreMergeCompatible, getParsedSeasonsForAnime } from './anime-merge';
import { toDbTimestamp } from './query-shared';
import { normalizeAnimeIdentityKey } from './storage';
import { nowMs } from './time';

export interface AnimeMergeRecommendation {
  recommendationId: number;
  animeIds: [number, number];
}

export interface AnimeConflictRecommendationOptions {
  survivor?: 'target' | 'existing';
  /** Automatic matches must be exact; manual assignment is authoritative. */
  matchConfidence?: 'exact' | 'weak' | 'manual';
}

interface AnimeTitleRow {
  canonical_title: string;
  title_romaji: string | null;
  title_english: string | null;
  title_native: string | null;
}

function getAnimeTitles(db: DatabaseSync, animeId: number): AnimeTitleRow | null {
  return db
    .prepare(
      `SELECT canonical_title, title_romaji, title_english, title_native
       FROM imm_anime
       WHERE anime_id = ?`,
    )
    .get(animeId) as AnimeTitleRow | null;
}

function getParsedTitles(db: DatabaseSync, animeId: number): Array<string | null> {
  return (
    db.prepare('SELECT parsed_title FROM imm_videos WHERE anime_id = ?').all(animeId) as Array<{
      parsed_title: string | null;
    }>
  ).map((row) => row.parsed_title);
}

function stripSeasonIdentitySuffix(title: string): string {
  return title
    .replace(/\bseason\s*\d{1,2}\b/gi, ' ')
    .replace(/\b\d{1,2}(?:st|nd|rd|th)\s+season\b/gi, ' ')
    .replace(/\bs\d{1,2}\b/gi, ' ');
}

export function hasExactStoredTitleMatch(
  db: DatabaseSync,
  targetAnimeId: number,
  conflictAnimeId: number,
): boolean {
  const target = getAnimeTitles(db, targetAnimeId);
  const conflict = getAnimeTitles(db, conflictAnimeId);
  if (!target || !conflict) return false;
  const targetKeys = [target.canonical_title, ...getParsedTitles(db, targetAnimeId)]
    .filter((title): title is string => Boolean(title?.trim()))
    .map((title) => normalizeAnimeIdentityKey(stripSeasonIdentitySuffix(title)))
    .filter(Boolean);
  const anilistTitleKeys = [
    conflict.title_romaji,
    conflict.title_english,
    conflict.title_native,
    conflict.canonical_title,
  ]
    .filter((title): title is string => Boolean(title?.trim()))
    .map(normalizeAnimeIdentityKey)
    .filter(Boolean);
  return targetKeys.some((key) => anilistTitleKeys.includes(key));
}

export function shouldRecommendAnilistConflict(
  db: DatabaseSync,
  targetAnimeId: number,
  conflictAnimeId: number,
  options: AnimeConflictRecommendationOptions,
): boolean {
  if (options.survivor === 'target' || options.matchConfidence === 'manual') return false;
  if (
    !animeSeasonsAreMergeCompatible(
      getParsedSeasonsForAnime(db, targetAnimeId),
      getParsedSeasonsForAnime(db, conflictAnimeId),
    )
  ) {
    return false;
  }
  return (
    options.matchConfidence === 'weak' ||
    (options.matchConfidence === undefined &&
      !hasExactStoredTitleMatch(db, targetAnimeId, conflictAnimeId))
  );
}

export function recordAnimeMergeRecommendation(
  db: DatabaseSync,
  firstCandidateAnimeId: number,
  secondCandidateAnimeId: number,
  anilistId: number,
): void {
  const firstAnimeId = Math.min(firstCandidateAnimeId, secondCandidateAnimeId);
  const secondAnimeId = Math.max(firstCandidateAnimeId, secondCandidateAnimeId);
  const timestamp = toDbTimestamp(nowMs());
  db.prepare(
    `INSERT INTO imm_anime_merge_recommendations(
       first_anime_id, second_anime_id, anilist_id, status, CREATED_DATE, LAST_UPDATE_DATE
     ) VALUES (?, ?, ?, 'pending', ?, ?)
     ON CONFLICT(first_anime_id, second_anime_id, anilist_id) DO UPDATE SET
       LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE`,
  ).run(firstAnimeId, secondAnimeId, anilistId, timestamp, timestamp);
}

export function hasDismissedAnimeMergeRecommendation(
  db: DatabaseSync,
  firstCandidateAnimeId: number,
  secondCandidateAnimeId: number,
): boolean {
  const firstAnimeId = Math.min(firstCandidateAnimeId, secondCandidateAnimeId);
  const secondAnimeId = Math.max(firstCandidateAnimeId, secondCandidateAnimeId);
  return Boolean(
    db
      .prepare(
        `SELECT 1
         FROM imm_anime_merge_recommendations
         WHERE first_anime_id = ?
           AND second_anime_id = ?
           AND status = 'dismissed'
         LIMIT 1`,
      )
      .get(firstAnimeId, secondAnimeId),
  );
}

export function getAnimeMergeRecommendations(db: DatabaseSync): AnimeMergeRecommendation[] {
  return (
    db
      .prepare(
        `SELECT recommendation_id AS recommendationId,
                first_anime_id AS firstAnimeId,
                second_anime_id AS secondAnimeId
         FROM imm_anime_merge_recommendations
         WHERE status = 'pending'
         ORDER BY recommendation_id ASC`,
      )
      .all() as Array<{
      recommendationId: number;
      firstAnimeId: number;
      secondAnimeId: number;
    }>
  ).map((row) => ({
    recommendationId: row.recommendationId,
    animeIds: [row.firstAnimeId, row.secondAnimeId],
  }));
}

export function dismissAnimeMergeRecommendation(
  db: DatabaseSync,
  recommendationId: number,
): boolean {
  const result = db
    .prepare(
      `UPDATE imm_anime_merge_recommendations
       SET status = 'dismissed', LAST_UPDATE_DATE = ?
       WHERE recommendation_id = ? AND status = 'pending'`,
    )
    .run(toDbTimestamp(nowMs()), recommendationId) as { changes: number };
  return result.changes > 0;
}
