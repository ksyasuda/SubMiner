import type { DatabaseSync } from './sqlite';
import type {
  KanjiAnimeAppearanceRow,
  KanjiDetailRow,
  KanjiOccurrenceRow,
  KanjiStatsRow,
  KanjiWordRow,
  SentenceSearchOptions,
  SentenceSearchResultRow,
  SessionEventRow,
  SimilarWordRow,
  StatsExcludedWordRow,
  VocabularyStatsRow,
  VocabularyStatsSummary,
  WordAnimeAppearanceRow,
  WordDetailRow,
  WordOccurrenceRow,
} from './types';
import { fromDbTimestamp, toDbTimestamp } from './query-shared';
import { nowMs } from './time';
import {
  areLexicalDailyRollupsReady,
  getLexicalDailyRollups,
  localEpochDaySql,
} from './lexical-rollups';
import { isVocabularyStatsRowVisible } from './vocabulary-visibility';

const VOCABULARY_STATS_FILTER_OVERSAMPLE_FACTOR = 4;
const VOCABULARY_STATS_FILTER_OVERSAMPLE_MIN = 100;
const VOCABULARY_CHART_LIMIT = 12;
const VOCABULARY_CHART_PAGE_SIZE = 100;
const EXCLUSION_ALIAS_BATCH_SIZE = 300;
const VOCABULARY_SUMMARY_SCAN_BATCH_SIZE = 5_000;
const SENTENCE_SEARCH_DEFAULT_LIMIT = 50;
const SENTENCE_SEARCH_MAX_LIMIT = 100;
const KANJI_PATTERN = /\p{Script=Han}/gu;

export interface VocabularyChartData {
  ready: boolean;
  topWords: Array<{ wordId: number; headword: string; frequency: number }>;
  topWordsWithoutNames: Array<{ wordId: number; headword: string; frequency: number }>;
  newWordsTimeline: Array<{ epochDay: number; wordCount: number }>;
  newWordsTimelineWithoutNames: Array<{ epochDay: number; wordCount: number }>;
}

function resolveSentenceSearchLimit(limit: number): number {
  if (!Number.isFinite(limit)) return SENTENCE_SEARCH_DEFAULT_LIMIT;
  const normalized = Math.floor(limit);
  if (normalized <= 0) return SENTENCE_SEARCH_DEFAULT_LIMIT;
  return Math.min(normalized, SENTENCE_SEARCH_MAX_LIMIT);
}

export function splitSentenceSearchTerms(query: string): string[] {
  return query
    .trim()
    .split(/\s+/)
    .map((term) => term.trim())
    .filter(Boolean)
    .slice(0, 8);
}

function escapeLikeTerm(term: string): string {
  return term.replace(/[\\%_]/g, (match) => `\\${match}`);
}

function uniqueNonEmptyTerms(values: readonly string[] | undefined): string[] {
  const seen = new Set<string>();
  const terms: string[] = [];
  for (const value of values ?? []) {
    const term = value.trim();
    if (!term || seen.has(term)) continue;
    seen.add(term);
    terms.push(term);
  }
  return terms;
}

function getHeadwordCandidatesForSentenceSearchTerm(
  term: string,
  options: SentenceSearchOptions | undefined,
): string[] {
  const headwords =
    options?.headwordTerms
      ?.filter((entry) => entry.term === term)
      .flatMap((entry) => entry.headwords) ?? [];
  return uniqueNonEmptyTerms(headwords);
}

function uniqueKanji(text: string): string[] {
  return Array.from(new Set(text.match(KANJI_PATTERN) ?? []));
}

export function getVocabularyStats(
  db: DatabaseSync,
  limit = 100,
  excludePos?: string[],
): VocabularyStatsRow[] {
  const pageSize = Math.max(
    limit,
    limit * VOCABULARY_STATS_FILTER_OVERSAMPLE_FACTOR,
    limit + VOCABULARY_STATS_FILTER_OVERSAMPLE_MIN,
  );
  const hasExclude = excludePos && excludePos.length > 0;
  const placeholders = hasExclude ? excludePos.map(() => '?').join(', ') : '';
  const whereClause = hasExclude
    ? `WHERE (part_of_speech IS NULL OR part_of_speech NOT IN (${placeholders}))`
    : '';
  // The page is selected before the join so `animeCount` is only computed for the
  // rows being returned. Aggregating first made every request walk each word's
  // entire occurrence history — seconds of blocked event loop on a large library,
  // because only the ordering, not the aggregate, decides which rows survive.
  const stmt = db.prepare(`
    WITH page AS (
      SELECT id, headword, word, reading, part_of_speech, pos1, pos2, pos3,
             frequency, frequency_rank, first_seen, last_seen
      FROM imm_words
      ${whereClause}
      ORDER BY frequency DESC, id
      LIMIT ? OFFSET ?
    )
    SELECT p.id AS wordId, p.headword, p.word, p.reading,
      p.part_of_speech AS partOfSpeech, p.pos1, p.pos2, p.pos3,
      p.frequency, p.frequency_rank AS frequencyRank,
      p.first_seen AS firstSeen, p.last_seen AS lastSeen,
      COUNT(DISTINCT sl.anime_id) AS animeCount
    FROM page p
    LEFT JOIN imm_word_line_occurrences o ON o.word_id = p.id
    LEFT JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id AND sl.anime_id IS NOT NULL
    GROUP BY p.id
    ORDER BY p.frequency DESC, p.id
  `);
  const visibleRows: VocabularyStatsRow[] = [];
  let offset = 0;

  while (visibleRows.length < limit) {
    const params = hasExclude ? [...excludePos, pageSize, offset] : [pageSize, offset];
    const page = stmt.all(...params) as VocabularyStatsRow[];
    if (page.length === 0) break;
    visibleRows.push(...page.filter(isVocabularyStatsRowVisible));
    offset += page.length;
  }

  return visibleRows.slice(0, limit);
}

/**
 * Chart data is intentionally independent of the paginated vocabulary tables.
 * Top words use the frequency index; new-word history reads permanent daily
 * lexical rollups rather than loading every vocabulary row into the dashboard.
 */
export function getVocabularyChartData(db: DatabaseSync): VocabularyChartData {
  const ready = areLexicalDailyRollupsReady(db);
  const excludedAliases = new Set(
    getStatsExcludedWords(db).flatMap((word) => excludedVocabularyAliases(word)),
  );
  const isExcluded = (word: Pick<VocabularyStatsRow, 'headword' | 'word' | 'reading'>): boolean =>
    excludedVocabularyAliases(word).some((alias) => excludedAliases.has(alias));
  const topWords = getTopVocabularyChartWords(db, isExcluded);
  const rollups = ready ? getLexicalDailyRollups(db) : [];
  const timeline = new Map(rollups.map((row) => [row.epochDay, { ...row }]));
  if (excludedAliases.size > 0 && ready) {
    const aliases = [...excludedAliases];
    const excludedRows = new Map<
      number,
      Pick<VocabularyStatsRow, 'headword' | 'word' | 'reading' | 'pos2'> & {
        wordId: number;
        epochDay: number;
      }
    >();
    for (let offset = 0; offset < aliases.length; offset += EXCLUSION_ALIAS_BATCH_SIZE) {
      const batch = aliases.slice(offset, offset + EXCLUSION_ALIAS_BATCH_SIZE);
      const placeholders = batch.map(() => '?').join(', ');
      const rows = db
        .prepare(
          `
            SELECT id AS wordId, headword, word, reading, pos2,
              ${localEpochDaySql('first_seen')} AS epochDay
            FROM imm_words
            WHERE vocabulary_visible = 1
              AND (headword IN (${placeholders}) OR word IN (${placeholders}) OR reading IN (${placeholders}))
          `,
        )
        .all(...batch, ...batch, ...batch) as Array<
        Pick<VocabularyStatsRow, 'headword' | 'word' | 'reading' | 'pos2'> & {
          wordId: number;
          epochDay: number;
        }
      >;
      for (const row of rows) excludedRows.set(row.wordId, row);
    }
    for (const word of excludedRows.values()) {
      if (!isExcluded(word)) continue;
      const rollup = timeline.get(word.epochDay);
      if (!rollup) continue;
      rollup.wordCount -= 1;
      if (word.pos2 !== '固有名詞') rollup.wordCountWithoutNames -= 1;
    }
  }
  return {
    ready,
    topWords: topWords.all.map((word) => ({
      wordId: word.wordId,
      headword: vocabularyDisplayHeadword(word),
      frequency: word.frequency,
    })),
    topWordsWithoutNames: topWords.withoutNames.map((word) => ({
      wordId: word.wordId,
      headword: vocabularyDisplayHeadword(word),
      frequency: word.frequency,
    })),
    newWordsTimeline: [...timeline.values()]
      .filter((row) => row.wordCount > 0)
      .map((row) => ({ epochDay: row.epochDay, wordCount: row.wordCount })),
    newWordsTimelineWithoutNames: [...timeline.values()]
      .filter((row) => row.wordCountWithoutNames > 0)
      .map((row) => ({ epochDay: row.epochDay, wordCount: row.wordCountWithoutNames })),
  };
}

function getTopVocabularyChartWords(
  db: DatabaseSync,
  isExcluded: (word: Pick<VocabularyStatsRow, 'headword' | 'word' | 'reading'>) => boolean,
): { all: VocabularyStatsRow[]; withoutNames: VocabularyStatsRow[] } {
  const stmt = db.prepare(`
    SELECT id AS wordId, headword, word, reading,
      part_of_speech AS partOfSpeech, pos1, pos2, pos3,
      frequency, frequency_rank AS frequencyRank,
      first_seen AS firstSeen, last_seen AS lastSeen,
      0 AS animeCount
    FROM imm_words
    ORDER BY frequency DESC, id
    LIMIT ? OFFSET ?
  `);
  const all: VocabularyStatsRow[] = [];
  const withoutNames: VocabularyStatsRow[] = [];
  let offset = 0;

  while (all.length < VOCABULARY_CHART_LIMIT || withoutNames.length < VOCABULARY_CHART_LIMIT) {
    const page = stmt.all(VOCABULARY_CHART_PAGE_SIZE, offset) as VocabularyStatsRow[];
    if (page.length === 0) break;
    for (const word of page) {
      if (!isVocabularyStatsRowVisible(word) || isExcluded(word)) continue;
      if (all.length < VOCABULARY_CHART_LIMIT) all.push(word);
      if (word.pos2 !== '固有名詞' && withoutNames.length < VOCABULARY_CHART_LIMIT) {
        withoutNames.push(word);
      }
    }
    offset += page.length;
  }

  return { all, withoutNames };
}

function excludedVocabularyAliases(
  word: Pick<VocabularyStatsRow, 'headword' | 'word' | 'reading'>,
): string[] {
  const aliases = [word.headword?.trim() ?? '', word.word?.trim() ?? ''].filter(Boolean);
  if (aliases.length === 0) aliases.push(word.reading?.trim() ?? '');
  return [...new Set(aliases)];
}

function vocabularyDisplayHeadword(
  word: Pick<VocabularyStatsRow, 'headword' | 'word' | 'reading'>,
): string {
  return word.headword?.trim() || word.word?.trim() || word.reading?.trim() || '';
}

function timestampSeconds(timestamp: number): number {
  return timestamp < 10_000_000_000 ? timestamp : Math.floor(timestamp / 1000);
}

export function getVocabularySummary(
  db: DatabaseSync,
  knownWords: ReadonlySet<string> | null,
  nowMs: number = Date.now(),
  scanBatchSize: number = VOCABULARY_SUMMARY_SCAN_BATCH_SIZE,
): VocabularyStatsSummary {
  // Visibility and exclusion rules live in JS, so rows are scanned in id-keyed
  // batches to keep memory bounded on large vocabularies.
  const scanStmt = db.prepare(`
    SELECT id AS wordId, headword, word, reading,
      part_of_speech AS partOfSpeech, pos1, pos2, pos3,
      frequency, frequency_rank AS frequencyRank,
      first_seen AS firstSeen, last_seen AS lastSeen,
      0 AS animeCount
    FROM imm_words
    WHERE id > ?
    ORDER BY id
    LIMIT ?
  `);
  const excludedAliases = new Set(
    getStatsExcludedWords(db).flatMap((word) => excludedVocabularyAliases(word)),
  );
  const weekAgoSec = nowMs / 1000 - 7 * 86_400;
  const summary: VocabularyStatsSummary = {
    uniqueWords: 0,
    uniqueWordsWithoutNames: 0,
    uniqueKanji: (db.prepare('SELECT COUNT(*) AS count FROM imm_kanji').get() as { count: number })
      .count,
    newThisWeek: 0,
    newThisWeekWithoutNames: 0,
    knownWordCount: knownWords ? 0 : null,
    knownWordCountWithoutNames: knownWords ? 0 : null,
  };

  let lastId = Number.MIN_SAFE_INTEGER;
  for (;;) {
    const words = scanStmt.all(lastId, scanBatchSize) as VocabularyStatsRow[];
    if (words.length === 0) break;
    lastId = words[words.length - 1]!.wordId;
    for (const word of words) {
      if (
        !isVocabularyStatsRowVisible(word) ||
        excludedVocabularyAliases(word).some((alias) => excludedAliases.has(alias))
      ) {
        continue;
      }
      const isName = word.pos2 === '固有名詞';
      const isNewThisWeek = timestampSeconds(fromDbTimestamp(word.firstSeen) ?? 0) >= weekAgoSec;
      const isKnown = knownWords?.has(vocabularyDisplayHeadword(word)) ?? false;
      summary.uniqueWords += 1;
      if (!isName) summary.uniqueWordsWithoutNames += 1;
      if (isNewThisWeek) {
        summary.newThisWeek += 1;
        if (!isName) summary.newThisWeekWithoutNames += 1;
      }
      if (isKnown) {
        summary.knownWordCount! += 1;
        if (!isName) summary.knownWordCountWithoutNames! += 1;
      }
    }
    if (words.length < scanBatchSize) break;
  }

  return summary;
}

export function getStatsExcludedWords(db: DatabaseSync): StatsExcludedWordRow[] {
  return db
    .prepare(
      `
        SELECT headword, word, reading
        FROM imm_stats_excluded_words
        ORDER BY headword COLLATE NOCASE, word COLLATE NOCASE, reading COLLATE NOCASE
      `,
    )
    .all() as StatsExcludedWordRow[];
}

export function replaceStatsExcludedWords(db: DatabaseSync, words: StatsExcludedWordRow[]): void {
  const now = toDbTimestamp(nowMs());
  const insertStmt = db.prepare(`
    INSERT OR IGNORE INTO imm_stats_excluded_words(
      headword,
      word,
      reading,
      CREATED_DATE,
      LAST_UPDATE_DATE
    )
    VALUES (?, ?, ?, ?, ?)
  `);

  db.exec('BEGIN IMMEDIATE');
  try {
    db.prepare('DELETE FROM imm_stats_excluded_words').run();
    for (const word of words) {
      insertStmt.run(word.headword, word.word, word.reading, now, now);
    }
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
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

export function searchSubtitleSentences(
  db: DatabaseSync,
  query: string,
  limit = SENTENCE_SEARCH_DEFAULT_LIMIT,
  options?: SentenceSearchOptions,
): SentenceSearchResultRow[] {
  const terms = splitSentenceSearchTerms(query);
  if (terms.length === 0) return [];
  const resolvedLimit = resolveSentenceSearchLimit(limit);

  const clauses: string[] = [];
  const params: string[] = [];
  for (const term of terms) {
    const likeTerm = `%${escapeLikeTerm(term)}%`;
    const headwords = getHeadwordCandidatesForSentenceSearchTerm(term, options);
    const headwordClause =
      headwords.length > 0
        ? `
        OR EXISTS (
          SELECT 1
          FROM imm_word_line_occurrences o
          JOIN imm_words w ON w.id = o.word_id
          WHERE o.line_id = l.line_id
            AND w.headword IN (${headwords.map(() => '?').join(', ')})
        )
      `
        : '';
    clauses.push(`
      (
        l.text LIKE ? ESCAPE '\\'
        OR v.canonical_title LIKE ? ESCAPE '\\'
        OR COALESCE(a.canonical_title, '') LIKE ? ESCAPE '\\'
        ${headwordClause}
      )
    `);
    params.push(likeTerm, likeTerm, likeTerm, ...headwords);
  }

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
          l.text AS text
        FROM imm_subtitle_lines l
        JOIN imm_videos v ON v.video_id = l.video_id
        LEFT JOIN imm_anime a ON a.anime_id = l.anime_id
        WHERE ${clauses.join(' AND ')}
        ORDER BY l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
      `,
    )
    .all(...params, resolvedLimit) as unknown as SentenceSearchResultRow[];
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
      FROM imm_session_events WHERE session_id = ? ORDER BY CAST(ts_ms AS REAL) ASC LIMIT ?
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
    ORDER BY CAST(ts_ms AS REAL) ASC
    LIMIT ?
  `);
  const rows = stmt.all(sessionId, ...eventTypes, limit) as Array<
    SessionEventRow & {
      tsMs: number | string;
    }
  >;
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

  const clauses: string[] = [];
  const params: string[] = [];
  const reading = word.reading.trim();
  if (reading !== '') {
    clauses.push('reading = ?');
    params.push(word.reading);
  }

  for (const kanji of uniqueKanji(word.headword)) {
    clauses.push("headword LIKE ? ESCAPE '\\'");
    params.push(`%${escapeLikeTerm(kanji)}%`);
  }

  if (clauses.length === 0) return [];

  const orderBy =
    reading !== '' ? 'CASE WHEN reading = ? THEN 0 ELSE 1 END, frequency DESC' : 'frequency DESC';
  const orderParams = reading !== '' ? [word.reading] : [];

  return db
    .prepare(
      `
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE id != ?
    AND (${clauses.join(' OR ')})
    ORDER BY ${orderBy}
    LIMIT ?
  `,
    )
    .all(wordId, ...params, ...orderParams, limit) as SimilarWordRow[];
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
