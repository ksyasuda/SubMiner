import { createHash } from 'node:crypto';
import type { DatabaseSync } from './sqlite';
import { buildCoverBlobReference, normalizeCoverBlobBytes } from './storage';
import { rebuildLifetimeSummaries, rebuildLifetimeSummariesInTransaction } from './lifetime';
import { getRollupGroupsForSessions, refreshRollupsForGroupsInTransaction } from './maintenance';
import { nowMs } from './time';
import { resolveAnimeAnilistConflict } from './anime-season-repair';
import { PartOfSpeech, type MergedToken } from '../../../types';
import { shouldExcludeTokenFromVocabularyPersistence } from '../tokenizer/annotation-stage';
import { deriveStoredPartOfSpeech } from '../tokenizer/part-of-speech';
import {
  applyLexicalRemovals,
  cleanupUnusedCoverArtBlobHash,
  deleteSessionsByIds,
  findSharedCoverBlobHash,
  planLexicalRemovalsForSessions,
  planLexicalRemovalsForVideos,
  toDbMs,
  toDbTimestamp,
} from './query-shared';

type CleanupVocabularyRow = {
  id: number;
  word: string;
  headword: string;
  reading: string | null;
  part_of_speech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
  first_seen: number | null;
  last_seen: number | null;
  frequency: number | null;
};

type ResolvedVocabularyPos = {
  headword: string;
  reading: string;
  hasPosMetadata: boolean;
  partOfSpeech: PartOfSpeech;
  pos1: string;
  pos2: string;
  pos3: string;
};

type CleanupVocabularyStatsOptions = {
  resolveLegacyPos?: (row: CleanupVocabularyRow) => Promise<{
    headword: string;
    reading: string;
    partOfSpeech: string;
    pos1: string;
    pos2: string;
    pos3: string;
  } | null>;
};

function toStoredWordToken(row: {
  word: string;
  headword: string;
  part_of_speech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
}): MergedToken {
  return {
    surface: row.word || row.headword || '',
    reading: '',
    headword: row.headword || row.word || '',
    startPos: 0,
    endPos: 0,
    partOfSpeech: deriveStoredPartOfSpeech({
      partOfSpeech: row.part_of_speech,
      pos1: row.pos1,
    }),
    pos1: row.pos1 ?? '',
    pos2: row.pos2 ?? '',
    pos3: row.pos3 ?? '',
    isMerged: true,
    isKnown: false,
    isNPlusOneTarget: false,
  };
}

function normalizePosField(value: string | null | undefined): string {
  return typeof value === 'string' ? value.trim() : '';
}

function resolveStoredVocabularyPos(row: CleanupVocabularyRow): ResolvedVocabularyPos | null {
  const headword = normalizePosField(row.headword);
  const reading = normalizePosField(row.reading);
  const partOfSpeechRaw = typeof row.part_of_speech === 'string' ? row.part_of_speech.trim() : '';
  const pos1 = normalizePosField(row.pos1);
  const pos2 = normalizePosField(row.pos2);
  const pos3 = normalizePosField(row.pos3);

  if (!headword && !reading && !partOfSpeechRaw && !pos1 && !pos2 && !pos3) {
    return null;
  }

  return {
    headword: headword || normalizePosField(row.word),
    reading,
    hasPosMetadata: Boolean(partOfSpeechRaw || pos1 || pos2 || pos3),
    partOfSpeech: deriveStoredPartOfSpeech({
      partOfSpeech: partOfSpeechRaw,
      pos1,
    }),
    pos1,
    pos2,
    pos3,
  };
}

function hasStructuredPos(pos: ResolvedVocabularyPos | null): boolean {
  return Boolean(pos?.hasPosMetadata && (pos.pos1 || pos.pos2 || pos.pos3 || pos.partOfSpeech));
}

function needsLegacyVocabularyMetadataRepair(
  row: CleanupVocabularyRow,
  stored: ResolvedVocabularyPos | null,
): boolean {
  if (!stored) {
    return true;
  }

  if (!hasStructuredPos(stored)) {
    return true;
  }

  if (!stored.reading) {
    return true;
  }

  if (!stored.headword) {
    return true;
  }

  return stored.headword === normalizePosField(row.word);
}

function shouldUpdateStoredVocabularyPos(
  row: CleanupVocabularyRow,
  next: ResolvedVocabularyPos,
): boolean {
  return (
    normalizePosField(row.headword) !== next.headword ||
    normalizePosField(row.reading) !== next.reading ||
    (next.hasPosMetadata &&
      (normalizePosField(row.part_of_speech) !== next.partOfSpeech ||
        normalizePosField(row.pos1) !== next.pos1 ||
        normalizePosField(row.pos2) !== next.pos2 ||
        normalizePosField(row.pos3) !== next.pos3))
  );
}

function chooseMergedPartOfSpeech(
  current: string | null | undefined,
  incoming: ResolvedVocabularyPos,
): string {
  const normalizedCurrent = normalizePosField(current);
  if (
    normalizedCurrent &&
    normalizedCurrent !== PartOfSpeech.other &&
    incoming.partOfSpeech === PartOfSpeech.other
  ) {
    return normalizedCurrent;
  }
  return incoming.partOfSpeech;
}

async function maybeResolveLegacyVocabularyPos(
  row: CleanupVocabularyRow,
  options: CleanupVocabularyStatsOptions,
): Promise<ResolvedVocabularyPos | null> {
  const stored = resolveStoredVocabularyPos(row);
  if (!needsLegacyVocabularyMetadataRepair(row, stored) || !options.resolveLegacyPos) {
    return stored;
  }

  const resolved = await options.resolveLegacyPos(row);
  if (resolved) {
    return {
      headword: normalizePosField(resolved.headword) || normalizePosField(row.word),
      reading: normalizePosField(resolved.reading),
      hasPosMetadata: true,
      partOfSpeech: deriveStoredPartOfSpeech({
        partOfSpeech: resolved.partOfSpeech,
        pos1: resolved.pos1,
      }),
      pos1: normalizePosField(resolved.pos1),
      pos2: normalizePosField(resolved.pos2),
      pos3: normalizePosField(resolved.pos3),
    };
  }

  return stored;
}

export async function cleanupVocabularyStats(
  db: DatabaseSync,
  options: CleanupVocabularyStatsOptions = {},
): Promise<{ scanned: number; kept: number; deleted: number; repaired: number }> {
  const rows = db
    .prepare(
      `SELECT id, word, headword, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
       FROM imm_words`,
    )
    .all() as CleanupVocabularyRow[];
  const findDuplicateStmt = db.prepare(
    `SELECT id, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
     FROM imm_words
     WHERE headword = ? AND word = ? AND reading = ? AND id != ?`,
  );
  const deleteStmt = db.prepare('DELETE FROM imm_words WHERE id = ?');
  const updateStmt = db.prepare(
    `UPDATE imm_words
     SET headword = ?, reading = ?, part_of_speech = ?, pos1 = ?, pos2 = ?, pos3 = ?
     WHERE id = ?`,
  );
  const mergeWordStmt = db.prepare(
    `UPDATE imm_words
     SET
       frequency = COALESCE(frequency, 0) + ?,
       part_of_speech = ?,
       pos1 = ?,
       pos2 = ?,
       pos3 = ?,
       first_seen = MIN(COALESCE(first_seen, ?), ?),
       last_seen = MAX(COALESCE(last_seen, ?), ?)
     WHERE id = ?`,
  );
  const moveOccurrencesStmt = db.prepare(
    `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count, seen_ms)
     SELECT line_id, ?, occurrence_count, seen_ms
     FROM imm_word_line_occurrences
     WHERE word_id = ?
     ON CONFLICT(line_id, word_id) DO UPDATE SET
       occurrence_count = imm_word_line_occurrences.occurrence_count + excluded.occurrence_count,
       seen_ms = COALESCE(imm_word_line_occurrences.seen_ms, excluded.seen_ms)`,
  );
  const deleteOccurrencesStmt = db.prepare(
    'DELETE FROM imm_word_line_occurrences WHERE word_id = ?',
  );
  let kept = 0;
  let deleted = 0;
  let repaired = 0;

  for (const row of rows) {
    const resolvedPos = await maybeResolveLegacyVocabularyPos(row, options);
    const shouldRepair = Boolean(resolvedPos && shouldUpdateStoredVocabularyPos(row, resolvedPos));
    if (resolvedPos && shouldRepair) {
      const duplicate = findDuplicateStmt.get(
        resolvedPos.headword,
        row.word,
        resolvedPos.reading,
        row.id,
      ) as {
        id: number;
        part_of_speech: string | null;
        pos1: string | null;
        pos2: string | null;
        pos3: string | null;
        first_seen: number | null;
        last_seen: number | null;
        frequency: number | null;
      } | null;
      if (duplicate) {
        moveOccurrencesStmt.run(duplicate.id, row.id);
        deleteOccurrencesStmt.run(row.id);
        mergeWordStmt.run(
          row.frequency ?? 0,
          chooseMergedPartOfSpeech(duplicate.part_of_speech, resolvedPos),
          normalizePosField(duplicate.pos1) || resolvedPos.pos1,
          normalizePosField(duplicate.pos2) || resolvedPos.pos2,
          normalizePosField(duplicate.pos3) || resolvedPos.pos3,
          row.first_seen ?? duplicate.first_seen ?? 0,
          row.first_seen ?? duplicate.first_seen ?? 0,
          row.last_seen ?? duplicate.last_seen ?? 0,
          row.last_seen ?? duplicate.last_seen ?? 0,
          duplicate.id,
        );
        deleteStmt.run(row.id);
        repaired += 1;
        deleted += 1;
        continue;
      }

      updateStmt.run(
        resolvedPos.headword,
        resolvedPos.reading,
        resolvedPos.partOfSpeech,
        resolvedPos.pos1,
        resolvedPos.pos2,
        resolvedPos.pos3,
        row.id,
      );
      repaired += 1;
    }

    const effectiveRow = {
      ...row,
      headword: resolvedPos?.headword ?? row.headword,
      reading: resolvedPos?.reading ?? row.reading,
      part_of_speech: resolvedPos?.hasPosMetadata ? resolvedPos.partOfSpeech : row.part_of_speech,
      pos1: resolvedPos?.pos1 ?? row.pos1,
      pos2: resolvedPos?.pos2 ?? row.pos2,
      pos3: resolvedPos?.pos3 ?? row.pos3,
    };
    const missingPos =
      !normalizePosField(effectiveRow.part_of_speech) &&
      !normalizePosField(effectiveRow.pos1) &&
      !normalizePosField(effectiveRow.pos2) &&
      !normalizePosField(effectiveRow.pos3);
    if (
      missingPos ||
      shouldExcludeTokenFromVocabularyPersistence(toStoredWordToken(effectiveRow))
    ) {
      deleteStmt.run(row.id);
      deleted += 1;
      continue;
    }
    kept += 1;
  }

  return {
    scanned: rows.length,
    kept,
    deleted,
    repaired,
  };
}

export function upsertCoverArt(
  db: DatabaseSync,
  videoId: number,
  art: {
    anilistId: number | null;
    coverUrl: string | null;
    coverBlob: ArrayBuffer | Uint8Array | Buffer | null;
    titleRomaji: string | null;
    titleEnglish: string | null;
    episodesTotal: number | null;
  },
): void {
  const existing = db
    .prepare(
      `
        SELECT cover_blob_hash AS coverBlobHash
        FROM imm_media_art
        WHERE video_id = ?
      `,
    )
    .get(videoId) as { coverBlobHash: string | null } | undefined;
  const sharedCoverBlobHash = findSharedCoverBlobHash(db, videoId, art.anilistId, art.coverUrl);
  const fetchedAtMs = toDbTimestamp(nowMs());
  const coverBlob = normalizeCoverBlobBytes(art.coverBlob);
  const computedCoverBlobHash =
    coverBlob && coverBlob.length > 0 ? createHash('sha256').update(coverBlob).digest('hex') : null;
  let coverBlobHash = computedCoverBlobHash ?? sharedCoverBlobHash ?? null;
  if (!coverBlobHash && (!coverBlob || coverBlob.length === 0)) {
    coverBlobHash = existing?.coverBlobHash ?? null;
  }

  if (computedCoverBlobHash && coverBlob && coverBlob.length > 0) {
    db.prepare(
      `
        INSERT INTO imm_cover_art_blobs (blob_hash, cover_blob, CREATED_DATE, LAST_UPDATE_DATE)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(blob_hash) DO UPDATE SET
          LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
      `,
    ).run(computedCoverBlobHash, coverBlob, fetchedAtMs, fetchedAtMs);
  }

  db.prepare(
    `
    INSERT INTO imm_media_art (
      video_id, anilist_id, cover_url, cover_blob, cover_blob_hash,
      title_romaji, title_english, episodes_total,
      fetched_at_ms, CREATED_DATE, LAST_UPDATE_DATE
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(video_id) DO UPDATE SET
      anilist_id = excluded.anilist_id,
      cover_url = excluded.cover_url,
      cover_blob = excluded.cover_blob,
      cover_blob_hash = excluded.cover_blob_hash,
      title_romaji = excluded.title_romaji,
      title_english = excluded.title_english,
      episodes_total = excluded.episodes_total,
      fetched_at_ms = excluded.fetched_at_ms,
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `,
  ).run(
    videoId,
    art.anilistId,
    art.coverUrl,
    coverBlobHash ? buildCoverBlobReference(coverBlobHash) : coverBlob,
    coverBlobHash,
    art.titleRomaji,
    art.titleEnglish,
    art.episodesTotal,
    fetchedAtMs,
    fetchedAtMs,
    fetchedAtMs,
  );

  if (existing?.coverBlobHash !== coverBlobHash) {
    cleanupUnusedCoverArtBlobHash(db, existing?.coverBlobHash ?? null);
  }
}

export function updateAnimeAnilistInfo(
  db: DatabaseSync,
  videoId: number,
  info: {
    anilistId: number;
    titleRomaji: string | null;
    titleEnglish: string | null;
    titleNative: string | null;
    episodesTotal: number | null;
    exactTitleMatch?: boolean;
  },
): void {
  const row = db.prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?').get(videoId) as {
    anime_id: number | null;
  } | null;
  if (!row?.anime_id) return;

  const repair = resolveAnimeAnilistConflict(db, row.anime_id, info.anilistId, {
    matchConfidence:
      info.exactTitleMatch === true ? 'exact' : info.exactTitleMatch === false ? 'weak' : undefined,
  });
  if (repair.mergeRecommended || repair.anilistAssignmentBlocked) return;
  const targetRow = db
    .prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?')
    .get(videoId) as {
    anime_id: number | null;
  } | null;
  if (!targetRow?.anime_id) return;

  db.prepare(
    `
    UPDATE imm_anime
    SET
      anilist_id = COALESCE(?, anilist_id),
      title_romaji = COALESCE(?, title_romaji),
      title_english = COALESCE(?, title_english),
      title_native = COALESCE(?, title_native),
      episodes_total = COALESCE(?, episodes_total),
      LAST_UPDATE_DATE = ?
    WHERE anime_id = ?
  `,
  ).run(
    info.anilistId,
    info.titleRomaji,
    info.titleEnglish,
    info.titleNative,
    info.episodesTotal,
    toDbTimestamp(nowMs()),
    targetRow.anime_id,
  );
  if (repair.movedVideos > 0 || repair.deletedAnimeRows > 0) {
    rebuildLifetimeSummaries(db);
  }
}

export function markVideoWatched(db: DatabaseSync, videoId: number, watched: boolean): void {
  db.prepare('UPDATE imm_videos SET watched = ?, LAST_UPDATE_DATE = ? WHERE video_id = ?').run(
    watched ? 1 : 0,
    toDbTimestamp(nowMs()),
    videoId,
  );
}

export function getVideoDurationMs(db: DatabaseSync, videoId: number): number {
  const row = db.prepare('SELECT duration_ms FROM imm_videos WHERE video_id = ?').get(videoId) as {
    duration_ms: number;
  } | null;
  return row?.duration_ms ?? 0;
}

export function isVideoWatched(db: DatabaseSync, videoId: number): boolean {
  const row = db.prepare('SELECT watched FROM imm_videos WHERE video_id = ?').get(videoId) as {
    watched: number;
  } | null;
  return row?.watched === 1;
}

export function deleteSession(db: DatabaseSync, sessionId: number): void {
  const sessionIds = [sessionId];

  db.exec('BEGIN IMMEDIATE');
  try {
    // Measured inside the write lock: the plan records what the delete removes,
    // and applying a plan taken against a different snapshot would subtract the
    // wrong totals from imm_words/imm_kanji.
    const lexicalRemovals = planLexicalRemovalsForSessions(db, sessionIds);
    const affectedRollupGroups = getRollupGroupsForSessions(db, sessionIds);
    deleteSessionsByIds(db, sessionIds);
    applyLexicalRemovals(db, lexicalRemovals);
    rebuildLifetimeSummariesInTransaction(db);
    refreshRollupsForGroupsInTransaction(db, affectedRollupGroups);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

export function deleteSessions(db: DatabaseSync, sessionIds: number[]): void {
  if (sessionIds.length === 0) return;

  db.exec('BEGIN IMMEDIATE');
  try {
    const lexicalRemovals = planLexicalRemovalsForSessions(db, sessionIds);
    const affectedRollupGroups = getRollupGroupsForSessions(db, sessionIds);
    deleteSessionsByIds(db, sessionIds);
    applyLexicalRemovals(db, lexicalRemovals);
    rebuildLifetimeSummariesInTransaction(db);
    refreshRollupsForGroupsInTransaction(db, affectedRollupGroups);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

/**
 * Delete an entire library entry: every episode of the anime, all of their
 * sessions and derived stats, and the anime row itself.
 *
 * Mirrors {@link deleteVideo} per episode, but batches the lexical refresh and
 * lifetime rebuild into a single transaction so a multi-episode title doesn't
 * pay for one full rebuild per episode.
 */
export function deleteAnime(db: DatabaseSync, animeId: number): void {
  db.exec('BEGIN IMMEDIATE');
  try {
    const videoIds = (
      db.prepare('SELECT video_id FROM imm_videos WHERE anime_id = ?').all(animeId) as Array<{
        video_id: number;
      }>
    ).map((row) => row.video_id);

    const lexicalRemovals = planLexicalRemovalsForVideos(db, videoIds);
    const coverBlobHashes: string[] = [];
    const sessionIds: number[] = [];
    for (const videoId of videoIds) {
      const artRow = db
        .prepare('SELECT cover_blob_hash AS coverBlobHash FROM imm_media_art WHERE video_id = ?')
        .get(videoId) as { coverBlobHash: string | null } | undefined;
      if (artRow?.coverBlobHash) {
        coverBlobHashes.push(artRow.coverBlobHash);
      }
      const sessions = db
        .prepare('SELECT session_id FROM imm_sessions WHERE video_id = ?')
        .all(videoId) as Array<{ session_id: number }>;
      sessionIds.push(...sessions.map((session) => session.session_id));
    }

    deleteSessionsByIds(db, sessionIds);
    const deleteLinesStmt = db.prepare('DELETE FROM imm_subtitle_lines WHERE video_id = ?');
    const deleteDailyStmt = db.prepare('DELETE FROM imm_daily_rollups WHERE video_id = ?');
    const deleteMonthlyStmt = db.prepare('DELETE FROM imm_monthly_rollups WHERE video_id = ?');
    const deleteArtStmt = db.prepare('DELETE FROM imm_media_art WHERE video_id = ?');
    const deleteVideoStmt = db.prepare('DELETE FROM imm_videos WHERE video_id = ?');
    for (const videoId of videoIds) {
      deleteLinesStmt.run(videoId);
      deleteDailyStmt.run(videoId);
      deleteMonthlyStmt.run(videoId);
      deleteArtStmt.run(videoId);
      deleteVideoStmt.run(videoId);
    }
    for (const coverBlobHash of new Set(coverBlobHashes)) {
      cleanupUnusedCoverArtBlobHash(db, coverBlobHash);
    }
    db.prepare('DELETE FROM imm_lifetime_anime WHERE anime_id = ?').run(animeId);
    db.prepare('DELETE FROM imm_anime WHERE anime_id = ?').run(animeId);
    applyLexicalRemovals(db, lexicalRemovals);
    rebuildLifetimeSummariesInTransaction(db);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

export function deleteVideo(db: DatabaseSync, videoId: number): void {
  db.exec('BEGIN IMMEDIATE');
  try {
    const artRow = db
      .prepare(
        `
        SELECT cover_blob_hash AS coverBlobHash
        FROM imm_media_art
        WHERE video_id = ?
      `,
      )
      .get(videoId) as { coverBlobHash: string | null } | undefined;
    const lexicalRemovals = planLexicalRemovalsForVideos(db, [videoId]);
    const sessions = db
      .prepare('SELECT session_id FROM imm_sessions WHERE video_id = ?')
      .all(videoId) as Array<{ session_id: number }>;

    deleteSessionsByIds(
      db,
      sessions.map((session) => session.session_id),
    );
    db.prepare('DELETE FROM imm_subtitle_lines WHERE video_id = ?').run(videoId);
    db.prepare('DELETE FROM imm_daily_rollups WHERE video_id = ?').run(videoId);
    db.prepare('DELETE FROM imm_monthly_rollups WHERE video_id = ?').run(videoId);
    db.prepare('DELETE FROM imm_media_art WHERE video_id = ?').run(videoId);
    cleanupUnusedCoverArtBlobHash(db, artRow?.coverBlobHash ?? null);
    db.prepare('DELETE FROM imm_videos WHERE video_id = ?').run(videoId);
    applyLexicalRemovals(db, lexicalRemovals);
    rebuildLifetimeSummariesInTransaction(db);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}
