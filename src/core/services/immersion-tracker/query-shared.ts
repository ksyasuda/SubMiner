import type { DatabaseSync } from './sqlite';

export const ACTIVE_SESSION_METRICS_CTE = `
  WITH active_session_metrics AS (
    SELECT
      t.session_id AS sessionId,
      MAX(t.total_watched_ms) AS totalWatchedMs,
      MAX(t.active_watched_ms) AS activeWatchedMs,
      MAX(t.lines_seen) AS linesSeen,
      MAX(t.tokens_seen) AS tokensSeen,
      MAX(t.cards_mined) AS cardsMined,
      MAX(t.lookup_count) AS lookupCount,
      MAX(t.lookup_hits) AS lookupHits,
      MAX(t.yomitan_lookup_count) AS yomitanLookupCount
    FROM imm_session_telemetry t
    JOIN imm_sessions s ON s.session_id = t.session_id
    WHERE s.ended_at_ms IS NULL
    GROUP BY t.session_id
  )
`;

export function makePlaceholders(values: number[]): string {
  return values.map(() => '?').join(',');
}

export function resolvedCoverBlobExpr(mediaAlias: string, blobStoreAlias: string): string {
  return `COALESCE(${blobStoreAlias}.cover_blob, CASE WHEN ${mediaAlias}.cover_blob_hash IS NULL THEN ${mediaAlias}.cover_blob ELSE NULL END)`;
}

export function cleanupUnusedCoverArtBlobHash(db: DatabaseSync, blobHash: string | null): void {
  if (!blobHash) {
    return;
  }
  db.prepare(
    `
      DELETE FROM imm_cover_art_blobs
      WHERE blob_hash = ?
        AND NOT EXISTS (
          SELECT 1
          FROM imm_media_art
          WHERE cover_blob_hash = ?
        )
    `,
  ).run(blobHash, blobHash);
}

export function findSharedCoverBlobHash(
  db: DatabaseSync,
  videoId: number,
  anilistId: number | null,
  coverUrl: string | null,
): string | null {
  if (anilistId !== null) {
    const byAnilist = db
      .prepare(
        `
          SELECT cover_blob_hash AS coverBlobHash
          FROM imm_media_art
          WHERE video_id != ?
            AND anilist_id = ?
            AND cover_blob_hash IS NOT NULL
          ORDER BY fetched_at_ms DESC, video_id DESC
          LIMIT 1
        `,
      )
      .get(videoId, anilistId) as { coverBlobHash: string | null } | undefined;
    if (byAnilist?.coverBlobHash) {
      return byAnilist.coverBlobHash;
    }
  }

  if (coverUrl) {
    const byUrl = db
      .prepare(
        `
          SELECT cover_blob_hash AS coverBlobHash
          FROM imm_media_art
          WHERE video_id != ?
            AND cover_url = ?
            AND cover_blob_hash IS NOT NULL
          ORDER BY fetched_at_ms DESC, video_id DESC
          LIMIT 1
        `,
      )
      .get(videoId, coverUrl) as { coverBlobHash: string | null } | undefined;
    return byUrl?.coverBlobHash ?? null;
  }

  return null;
}

type LexicalEntity = 'word' | 'kanji';

function getAffectedIdsForSessions(
  db: DatabaseSync,
  entity: LexicalEntity,
  sessionIds: number[],
): number[] {
  if (sessionIds.length === 0) return [];
  const table = entity === 'word' ? 'imm_word_line_occurrences' : 'imm_kanji_line_occurrences';
  const col = `${entity}_id`;
  return (
    db
      .prepare(
        `SELECT DISTINCT o.${col} AS id
         FROM ${table} o
         JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
         WHERE sl.session_id IN (${makePlaceholders(sessionIds)})`,
      )
      .all(...sessionIds) as Array<{ id: number }>
  ).map((row) => row.id);
}

function getAffectedIdsForVideo(
  db: DatabaseSync,
  entity: LexicalEntity,
  videoId: number,
): number[] {
  const table = entity === 'word' ? 'imm_word_line_occurrences' : 'imm_kanji_line_occurrences';
  const col = `${entity}_id`;
  return (
    db
      .prepare(
        `SELECT DISTINCT o.${col} AS id
         FROM ${table} o
         JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
         WHERE sl.video_id = ?`,
      )
      .all(videoId) as Array<{ id: number }>
  ).map((row) => row.id);
}

export function getAffectedWordIdsForSessions(db: DatabaseSync, sessionIds: number[]): number[] {
  return getAffectedIdsForSessions(db, 'word', sessionIds);
}

export function getAffectedKanjiIdsForSessions(db: DatabaseSync, sessionIds: number[]): number[] {
  return getAffectedIdsForSessions(db, 'kanji', sessionIds);
}

export function getAffectedWordIdsForVideo(db: DatabaseSync, videoId: number): number[] {
  return getAffectedIdsForVideo(db, 'word', videoId);
}

export function getAffectedKanjiIdsForVideo(db: DatabaseSync, videoId: number): number[] {
  return getAffectedIdsForVideo(db, 'kanji', videoId);
}

function refreshWordAggregates(db: DatabaseSync, wordIds: number[]): void {
  if (wordIds.length === 0) {
    return;
  }

  const rows = db
    .prepare(
      `
        SELECT
          w.id AS wordId,
          COALESCE(SUM(o.occurrence_count), 0) AS frequency,
          MIN(COALESCE(sl.CREATED_DATE, sl.LAST_UPDATE_DATE)) AS firstSeen,
          MAX(COALESCE(sl.LAST_UPDATE_DATE, sl.CREATED_DATE)) AS lastSeen
        FROM imm_words w
        LEFT JOIN imm_word_line_occurrences o ON o.word_id = w.id
        LEFT JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE w.id IN (${makePlaceholders(wordIds)})
        GROUP BY w.id
      `,
    )
    .all(...wordIds) as Array<{
    wordId: number;
    frequency: number;
    firstSeen: number | null;
    lastSeen: number | null;
  }>;
  const updateStmt = db.prepare(
    `
      UPDATE imm_words
      SET frequency = ?, first_seen = ?, last_seen = ?
      WHERE id = ?
    `,
  );
  const deleteStmt = db.prepare('DELETE FROM imm_words WHERE id = ?');

  for (const row of rows) {
    if (row.frequency <= 0 || row.firstSeen === null || row.lastSeen === null) {
      deleteStmt.run(row.wordId);
      continue;
    }
    updateStmt.run(
      row.frequency,
      Math.floor(row.firstSeen / 1000),
      Math.floor(row.lastSeen / 1000),
      row.wordId,
    );
  }
}

function refreshKanjiAggregates(db: DatabaseSync, kanjiIds: number[]): void {
  if (kanjiIds.length === 0) {
    return;
  }

  const rows = db
    .prepare(
      `
        SELECT
          k.id AS kanjiId,
          COALESCE(SUM(o.occurrence_count), 0) AS frequency,
          MIN(COALESCE(sl.CREATED_DATE, sl.LAST_UPDATE_DATE)) AS firstSeen,
          MAX(COALESCE(sl.LAST_UPDATE_DATE, sl.CREATED_DATE)) AS lastSeen
        FROM imm_kanji k
        LEFT JOIN imm_kanji_line_occurrences o ON o.kanji_id = k.id
        LEFT JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE k.id IN (${makePlaceholders(kanjiIds)})
        GROUP BY k.id
      `,
    )
    .all(...kanjiIds) as Array<{
    kanjiId: number;
    frequency: number;
    firstSeen: number | null;
    lastSeen: number | null;
  }>;
  const updateStmt = db.prepare(
    `
      UPDATE imm_kanji
      SET frequency = ?, first_seen = ?, last_seen = ?
      WHERE id = ?
    `,
  );
  const deleteStmt = db.prepare('DELETE FROM imm_kanji WHERE id = ?');

  for (const row of rows) {
    if (row.frequency <= 0 || row.firstSeen === null || row.lastSeen === null) {
      deleteStmt.run(row.kanjiId);
      continue;
    }
    updateStmt.run(
      row.frequency,
      Math.floor(row.firstSeen / 1000),
      Math.floor(row.lastSeen / 1000),
      row.kanjiId,
    );
  }
}

export function refreshLexicalAggregates(
  db: DatabaseSync,
  wordIds: number[],
  kanjiIds: number[],
): void {
  refreshWordAggregates(db, [...new Set(wordIds)]);
  refreshKanjiAggregates(db, [...new Set(kanjiIds)]);
}

export function deleteSessionsByIds(db: DatabaseSync, sessionIds: number[]): void {
  if (sessionIds.length === 0) {
    return;
  }

  const placeholders = makePlaceholders(sessionIds);
  db.prepare(`DELETE FROM imm_subtitle_lines WHERE session_id IN (${placeholders})`).run(
    ...sessionIds,
  );
  db.prepare(`DELETE FROM imm_session_telemetry WHERE session_id IN (${placeholders})`).run(
    ...sessionIds,
  );
  db.prepare(`DELETE FROM imm_session_events WHERE session_id IN (${placeholders})`).run(
    ...sessionIds,
  );
  db.prepare(`DELETE FROM imm_sessions WHERE session_id IN (${placeholders})`).run(...sessionIds);
}
