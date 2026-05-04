import type { DatabaseSync } from './sqlite';
import { SUBTITLE_ANNOTATION_EXCLUDED_TERMS } from '../tokenizer/subtitle-annotation-filter';
import { nowMs } from './time';

function quoteSqlString(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}

const SQL_EXCLUDED_VOCABULARY_TERMS = [...SUBTITLE_ANNOTATION_EXCLUDED_TERMS].map(quoteSqlString);
const SQL_EXCLUDED_VOCABULARY_TERMS_LIST =
  SQL_EXCLUDED_VOCABULARY_TERMS.length > 0 ? SQL_EXCLUDED_VOCABULARY_TERMS.join(', ') : "''";

export function visibleWordSql(wordAlias: string): string {
  return `(
    TRIM(COALESCE(${wordAlias}.word, '')) NOT IN (${SQL_EXCLUDED_VOCABULARY_TERMS_LIST})
    AND TRIM(COALESCE(${wordAlias}.headword, '')) NOT IN (${SQL_EXCLUDED_VOCABULARY_TERMS_LIST})
    AND TRIM(COALESCE(${wordAlias}.reading, '')) NOT IN (${SQL_EXCLUDED_VOCABULARY_TERMS_LIST})
  )`;
}

export function filteredWordOccurrenceCountSql(occurrenceAlias: string, wordAlias: string): string {
  return `CASE
    WHEN ${occurrenceAlias}.word_id IS NOT NULL AND ${visibleWordSql(wordAlias)}
      THEN ${occurrenceAlias}.occurrence_count
    ELSE 0
  END`;
}

export const SESSION_WORD_COUNTS_SELECT = `
  SELECT
    sl.session_id AS sessionId,
    COUNT(DISTINCT sl.line_id) AS persistedLineCount,
    COALESCE(SUM(${filteredWordOccurrenceCountSql('wlo', 'w')}), 0) AS filteredWordsSeen
  FROM imm_subtitle_lines sl
  LEFT JOIN imm_word_line_occurrences wlo ON wlo.line_id = sl.line_id
  LEFT JOIN imm_words w ON w.id = wlo.word_id
  GROUP BY sl.session_id
`;

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
  ),
  session_word_counts AS (
    ${SESSION_WORD_COUNTS_SELECT}
  )
`;

export const SESSION_WORD_COUNTS_CTE = `
  WITH session_word_counts AS (
    ${SESSION_WORD_COUNTS_SELECT}
  )
`;

export function sessionDisplayWordsExpr(
  sessionAlias: string,
  wordCountAlias: string,
  rawTokensExpr = `${sessionAlias}.tokens_seen`,
): string {
  return `CASE
    WHEN COALESCE(${wordCountAlias}.persistedLineCount, 0) > 0 THEN COALESCE(${wordCountAlias}.filteredWordsSeen, 0)
    ELSE COALESCE(${rawTokensExpr}, 0)
  END`;
}

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

export function toDbMs(ms: number | bigint): bigint {
  if (typeof ms === 'bigint') {
    return ms;
  }
  if (!Number.isFinite(ms)) {
    throw new TypeError(`Invalid database timestamp: ${ms}`);
  }
  return BigInt(Math.trunc(ms));
}

function normalizeTimestampString(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    throw new TypeError(`Invalid database timestamp: ${value}`);
  }

  const integerLike = /^(-?)(\d+)(?:\.0+)?$/.exec(trimmed);
  if (integerLike) {
    const sign = integerLike[1] ?? '';
    const digits = (integerLike[2] ?? '0').replace(/^0+(?=\d)/, '');
    return `${sign}${digits || '0'}`;
  }

  const parsed = Number(trimmed);
  if (!Number.isFinite(parsed)) {
    throw new TypeError(`Invalid database timestamp: ${value}`);
  }
  return JSON.stringify(Math.trunc(parsed));
}

export function toDbTimestamp(ms: number | bigint | string): string {
  const normalizeParsed = (parsed: number): string => JSON.stringify(Math.trunc(parsed));

  if (typeof ms === 'bigint') {
    return ms.toString();
  }
  if (typeof ms === 'string') {
    return normalizeTimestampString(ms);
  }
  if (!Number.isFinite(ms)) {
    throw new TypeError(`Invalid database timestamp: ${ms}`);
  }
  return normalizeParsed(ms);
}

export function currentDbTimestamp(): string {
  const testNowMs = globalThis.__subminerTestNowMs;
  if (typeof testNowMs === 'string') {
    return normalizeTimestampString(testNowMs);
  }
  if (typeof testNowMs === 'number' && Number.isFinite(testNowMs)) {
    return toDbTimestamp(testNowMs);
  }
  return toDbTimestamp(nowMs());
}

export function subtractDbTimestamp(
  timestampMs: number | bigint | string,
  deltaMs: number | bigint,
): string {
  return (BigInt(toDbTimestamp(timestampMs)) - BigInt(deltaMs)).toString();
}

export function fromDbTimestamp(ms: number | bigint | string | null | undefined): number | null {
  if (ms === null || ms === undefined) {
    return null;
  }
  if (typeof ms === 'number') {
    return ms;
  }
  if (typeof ms === 'bigint') {
    return Number(ms);
  }
  const normalized = normalizeTimestampString(ms);
  if (/^-?\d+$/.test(normalized)) {
    return Number(BigInt(normalized));
  }
  return Math.trunc(Number.parseFloat(normalized));
}

function getNumericCalendarValue(
  db: DatabaseSync,
  sql: string,
  timestampMs: number | bigint | string,
): number {
  const row = db.prepare(sql).get(toDbTimestamp(timestampMs)) as
    | { value: number | string | null }
    | undefined;
  return Number(row?.value ?? 0);
}

export function getLocalEpochDay(db: DatabaseSync, timestampMs: number | bigint | string): number {
  return getNumericCalendarValue(
    db,
    `
      SELECT CAST(
        julianday(CAST(? AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5
        AS INTEGER
      ) AS value
    `,
    timestampMs,
  );
}

export function getLocalMonthKey(db: DatabaseSync, timestampMs: number | bigint | string): number {
  return getNumericCalendarValue(
    db,
    `
      SELECT CAST(
        strftime('%Y%m', CAST(? AS REAL) / 1000, 'unixepoch', 'localtime')
        AS INTEGER
      ) AS value
    `,
    timestampMs,
  );
}

export function getLocalDayOfWeek(db: DatabaseSync, timestampMs: number | bigint | string): number {
  return getNumericCalendarValue(
    db,
    `
      SELECT CAST(
        strftime('%w', CAST(? AS REAL) / 1000, 'unixepoch', 'localtime')
        AS INTEGER
      ) AS value
    `,
    timestampMs,
  );
}

export function getLocalHourOfDay(db: DatabaseSync, timestampMs: number | bigint | string): number {
  return getNumericCalendarValue(
    db,
    `
      SELECT CAST(
        strftime('%H', CAST(? AS REAL) / 1000, 'unixepoch', 'localtime')
        AS INTEGER
      ) AS value
    `,
    timestampMs,
  );
}

export function getStartOfLocalDaySec(
  db: DatabaseSync,
  timestampMs: number | bigint | string,
): number {
  return getNumericCalendarValue(
    db,
    `
      SELECT CAST(
        strftime(
          '%s',
          CAST(? AS REAL) / 1000,
          'unixepoch',
          'localtime',
          'start of day',
          'utc'
        ) AS INTEGER
      ) AS value
    `,
    timestampMs,
  );
}

export function getStartOfLocalDayTimestamp(
  db: DatabaseSync,
  timestampMs: number | bigint | string,
): string {
  return `${getStartOfLocalDaySec(db, timestampMs)}000`;
}

export function getShiftedLocalDayTimestamp(
  db: DatabaseSync,
  timestampMs: number | bigint | string,
  dayOffset: number,
): string {
  const normalizedDayOffset = Math.trunc(dayOffset);
  const modifier =
    normalizedDayOffset >= 0 ? `+${normalizedDayOffset} days` : `${normalizedDayOffset} days`;
  const row = db
    .prepare(
      `
        SELECT strftime(
          '%s',
          CAST(? AS REAL) / 1000,
          'unixepoch',
          'localtime',
          'start of day',
          '${modifier}',
          'utc'
        ) AS value
      `,
    )
    .get(toDbTimestamp(timestampMs)) as { value: string | number | null } | undefined;
  return `${row?.value ?? '0'}000`;
}

export function getShiftedLocalDaySec(
  db: DatabaseSync,
  timestampMs: number | bigint | string,
  dayOffset: number,
): number {
  return Number(BigInt(getShiftedLocalDayTimestamp(db, timestampMs, dayOffset)) / 1000n);
}

export function getStartOfLocalDayMs(
  db: DatabaseSync,
  timestampMs: number | bigint | string,
): number {
  return getStartOfLocalDaySec(db, timestampMs) * 1000;
}
