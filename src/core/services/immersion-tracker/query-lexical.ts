import type { DatabaseSync } from './sqlite';
import type {
  KanjiAnimeAppearanceRow,
  KanjiDetailRow,
  KanjiOccurrenceRow,
  KanjiStatsRow,
  KanjiWordRow,
  SessionEventRow,
  SimilarWordRow,
  VocabularyStatsRow,
  WordAnimeAppearanceRow,
  WordDetailRow,
  WordOccurrenceRow,
} from './types';
import { fromDbTimestamp } from './query-shared';

export function getVocabularyStats(
  db: DatabaseSync,
  limit = 100,
  excludePos?: string[],
): VocabularyStatsRow[] {
  const hasExclude = excludePos && excludePos.length > 0;
  const placeholders = hasExclude ? excludePos.map(() => '?').join(', ') : '';
  const whereClause = hasExclude
    ? `WHERE (part_of_speech IS NULL OR part_of_speech NOT IN (${placeholders}))`
    : '';
  const stmt = db.prepare(`
    SELECT w.id AS wordId, w.headword, w.word, w.reading,
      w.part_of_speech AS partOfSpeech, w.pos1, w.pos2, w.pos3,
      w.frequency, w.frequency_rank AS frequencyRank,
      w.first_seen AS firstSeen, w.last_seen AS lastSeen,
      COUNT(DISTINCT sl.anime_id) AS animeCount
    FROM imm_words w
    LEFT JOIN imm_word_line_occurrences o ON o.word_id = w.id
    LEFT JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id AND sl.anime_id IS NOT NULL
    ${whereClause ? whereClause.replace('part_of_speech', 'w.part_of_speech') : ''}
    GROUP BY w.id
    ORDER BY w.frequency DESC LIMIT ?
  `);
  const params = hasExclude ? [...excludePos, limit] : [limit];
  return stmt.all(...params) as VocabularyStatsRow[];
}

export function getKanjiStats(db: DatabaseSync, limit = 100): KanjiStatsRow[] {
  const stmt = db.prepare(`
    SELECT id AS kanjiId, kanji, frequency,
      first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_kanji ORDER BY frequency DESC LIMIT ?
  `);
  return stmt.all(limit) as KanjiStatsRow[];
}

export function getWordOccurrences(
  db: DatabaseSync,
  headword: string,
  word: string,
  reading: string,
  limit = 100,
  offset = 0,
): WordOccurrenceRow[] {
  return db
    .prepare(
      `
        SELECT
          l.anime_id AS animeId,
          a.canonical_title AS animeTitle,
          l.video_id AS videoId,
          v.canonical_title AS videoTitle,
          v.source_path AS sourcePath,
          l.secondary_text AS secondaryText,
          l.session_id AS sessionId,
          l.line_index AS lineIndex,
          l.segment_start_ms AS segmentStartMs,
          l.segment_end_ms AS segmentEndMs,
          l.text AS text,
          o.occurrence_count AS occurrenceCount
        FROM imm_word_line_occurrences o
        JOIN imm_words w ON w.id = o.word_id
        JOIN imm_subtitle_lines l ON l.line_id = o.line_id
        JOIN imm_videos v ON v.video_id = l.video_id
        LEFT JOIN imm_anime a ON a.anime_id = l.anime_id
        WHERE w.headword = ? AND w.word = ? AND w.reading = ?
        ORDER BY l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
        OFFSET ?
      `,
    )
    .all(headword, word, reading, limit, offset) as unknown as WordOccurrenceRow[];
}

export function getKanjiOccurrences(
  db: DatabaseSync,
  kanji: string,
  limit = 100,
  offset = 0,
): KanjiOccurrenceRow[] {
  return db
    .prepare(
      `
        SELECT
          l.anime_id AS animeId,
          a.canonical_title AS animeTitle,
          l.video_id AS videoId,
          v.canonical_title AS videoTitle,
          v.source_path AS sourcePath,
          l.secondary_text AS secondaryText,
          l.session_id AS sessionId,
          l.line_index AS lineIndex,
          l.segment_start_ms AS segmentStartMs,
          l.segment_end_ms AS segmentEndMs,
          l.text AS text,
          o.occurrence_count AS occurrenceCount
        FROM imm_kanji_line_occurrences o
        JOIN imm_kanji k ON k.id = o.kanji_id
        JOIN imm_subtitle_lines l ON l.line_id = o.line_id
        JOIN imm_videos v ON v.video_id = l.video_id
        LEFT JOIN imm_anime a ON a.anime_id = l.anime_id
        WHERE k.kanji = ?
        ORDER BY l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
        OFFSET ?
      `,
    )
    .all(kanji, limit, offset) as unknown as KanjiOccurrenceRow[];
}

export function getSessionEvents(
  db: DatabaseSync,
  sessionId: number,
  limit = 500,
  eventTypes?: number[],
): SessionEventRow[] {
  if (!eventTypes || eventTypes.length === 0) {
    const stmt = db.prepare(`
      SELECT event_type AS eventType, ts_ms AS tsMs, payload_json AS payload
      FROM imm_session_events WHERE session_id = ? ORDER BY ts_ms ASC LIMIT ?
    `);
    const rows = stmt.all(sessionId, limit) as Array<SessionEventRow & { tsMs: number | string }>;
    return rows.map((row) => ({
      ...row,
      tsMs: fromDbTimestamp(row.tsMs) ?? 0,
    }));
  }

  const placeholders = eventTypes.map(() => '?').join(', ');
  const stmt = db.prepare(`
    SELECT event_type AS eventType, ts_ms AS tsMs, payload_json AS payload
    FROM imm_session_events
    WHERE session_id = ? AND event_type IN (${placeholders})
    ORDER BY ts_ms ASC
    LIMIT ?
  `);
  const rows = stmt.all(sessionId, ...eventTypes, limit) as Array<SessionEventRow & {
    tsMs: number | string;
  }>;
  return rows.map((row) => ({
    ...row,
    tsMs: fromDbTimestamp(row.tsMs) ?? 0,
  }));
}

export function getWordDetail(db: DatabaseSync, wordId: number): WordDetailRow | null {
  return db
    .prepare(
      `
    SELECT id AS wordId, headword, word, reading,
           part_of_speech AS partOfSpeech, pos1, pos2, pos3,
           frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_words WHERE id = ?
  `,
    )
    .get(wordId) as WordDetailRow | null;
}

export function getWordAnimeAppearances(
  db: DatabaseSync,
  wordId: number,
): WordAnimeAppearanceRow[] {
  return db
    .prepare(
      `
    SELECT a.anime_id AS animeId, a.canonical_title AS animeTitle,
           SUM(o.occurrence_count) AS occurrenceCount
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE o.word_id = ? AND sl.anime_id IS NOT NULL
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `,
    )
    .all(wordId) as WordAnimeAppearanceRow[];
}

export function getSimilarWords(db: DatabaseSync, wordId: number, limit = 10): SimilarWordRow[] {
  const word = db.prepare('SELECT headword, reading FROM imm_words WHERE id = ?').get(wordId) as {
    headword: string;
    reading: string;
  } | null;
  if (!word || word.headword.trim() === '') return [];
  return db
    .prepare(
      `
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE id != ?
    AND (reading = ? OR headword LIKE ? OR headword LIKE ?)
    ORDER BY frequency DESC
    LIMIT ?
  `,
    )
    .all(
      wordId,
      word.reading,
      `%${word.headword.charAt(0)}%`,
      `%${word.headword.charAt(word.headword.length - 1)}%`,
      limit,
    ) as SimilarWordRow[];
}

export function getKanjiDetail(db: DatabaseSync, kanjiId: number): KanjiDetailRow | null {
  return db
    .prepare(
      `
    SELECT id AS kanjiId, kanji, frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_kanji WHERE id = ?
  `,
    )
    .get(kanjiId) as KanjiDetailRow | null;
}

export function getKanjiAnimeAppearances(
  db: DatabaseSync,
  kanjiId: number,
): KanjiAnimeAppearanceRow[] {
  return db
    .prepare(
      `
    SELECT a.anime_id AS animeId, a.canonical_title AS animeTitle,
           SUM(o.occurrence_count) AS occurrenceCount
    FROM imm_kanji_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE o.kanji_id = ? AND sl.anime_id IS NOT NULL
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `,
    )
    .all(kanjiId) as KanjiAnimeAppearanceRow[];
}

export function getKanjiWords(db: DatabaseSync, kanjiId: number, limit = 20): KanjiWordRow[] {
  const kanjiRow = db.prepare('SELECT kanji FROM imm_kanji WHERE id = ?').get(kanjiId) as {
    kanji: string;
  } | null;
  if (!kanjiRow) return [];
  return db
    .prepare(
      `
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE headword LIKE ?
    ORDER BY frequency DESC
    LIMIT ?
  `,
    )
    .all(`%${kanjiRow.kanji}%`, limit) as KanjiWordRow[];
}
