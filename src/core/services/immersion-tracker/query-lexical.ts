import type { TrendFrequencyBand } from '../../../types/stats-wire';
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
import { forEachIdChunk, fromDbTimestamp, makePlaceholders, toDbTimestamp } from './query-shared';
import { nowMs } from './time';
import { narrowFullwidthKana, widenHalfwidthKana } from '../../utils/halfwidth-kana';
import { areLexicalDailyRollupsReady, localEpochDaySql } from './lexical-rollups';
import { isVocabularyStatsRowVisible } from './vocabulary-visibility';

// Merged tokens join MeCab tags, so a name can be stored as 固有名詞|一般 or 固有名詞|接尾.
const PROPER_NOUN_TAG = '固有名詞';
const NOT_A_NAME_SQL = `COALESCE(pos2, '') NOT LIKE '%${PROPER_NOUN_TAG}%'`;

/** The dictionary word an `imm_words` row belongs to. Legacy rows can lack a headword or form. */
export const DICTIONARY_WORD_KEY_SQL = `COALESCE(NULLIF(TRIM(headword), ''), NULLIF(TRIM(word), ''), TRIM(reading))`;

/** True for lines from a stream: only local videos have a path to cut clips from. */
export const NO_LOCAL_FILE_SQL = `TRIM(COALESCE(v.source_path, '')) = ''`;

// Sorts lines that can be mined ahead of those that can't, because newest-first alone
// lets a run of recent streams bury every line with a usable clip.
// Mirrors getStatsMineCardUnavailableReason in the stats client.
const MINEABLE_LINES_FIRST_SQL = `
  CASE
    WHEN NOT (${NO_LOCAL_FILE_SQL})
      AND l.segment_start_ms IS NOT NULL
      AND l.segment_end_ms > l.segment_start_ms
    THEN 0
    ELSE 1
  END`;

export function isProperNounPos2(pos2: string | null | undefined): boolean {
  return pos2?.includes(PROPER_NOUN_TAG) ?? false;
}

export interface NewWordFirstSeenRow {
  firstSeenDay: number;
  firstSeenDayWithoutNames: number | null;
  band: TrendFrequencyBand;
  wordCount: number;
}

const VOCABULARY_STATS_FILTER_OVERSAMPLE_FACTOR = 4;
const VOCABULARY_STATS_FILTER_OVERSAMPLE_MIN = 100;
const EXCLUSION_ALIAS_BATCH_SIZE = 300;
const VOCABULARY_SUMMARY_SCAN_BATCH_SIZE = 5_000;
const SENTENCE_SEARCH_DEFAULT_LIMIT = 50;
const SENTENCE_SEARCH_MAX_LIMIT = 100;
const KANJI_PATTERN = /\p{Script=Han}/gu;

export interface VocabularyChartData {
  ready: boolean;
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
  // Rows are paged and filtered before `animeCount` is computed, and only the
  // rows being returned are counted. The count walks each word's entire
  // occurrence history, and the most frequent words hold most occurrences, so
  // counting the oversampled page or aggregating first cost seconds of blocked
  // event loop on a large library.
  const stmt = db.prepare(`
    SELECT id AS wordId, headword, word, reading,
      part_of_speech AS partOfSpeech, pos1, pos2, pos3,
      frequency, frequency_rank AS frequencyRank,
      first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_words
    ${whereClause}
    ORDER BY frequency DESC, id
    LIMIT ? OFFSET ?
  `);
  const visibleRows: Array<Omit<VocabularyStatsRow, 'animeCount'>> = [];
  let offset = 0;

  while (visibleRows.length < limit) {
    const params = hasExclude ? [...excludePos, pageSize, offset] : [pageSize, offset];
    const page = stmt.all(...params) as Array<Omit<VocabularyStatsRow, 'animeCount'>>;
    if (page.length === 0) break;
    visibleRows.push(...page.filter(isVocabularyStatsRowVisible));
    offset += page.length;
  }

  const rows = visibleRows.slice(0, limit);
  const animeCounts = getWordAnimeCounts(
    db,
    rows.map((row) => row.wordId),
  );
  return rows.map((row) => ({ ...row, animeCount: animeCounts.get(row.wordId) ?? 0 }));
}

/** Number of distinct anime each word appears in, keyed by word id. */
function getWordAnimeCounts(db: DatabaseSync, wordIds: number[]): Map<number, number> {
  const counts = new Map<number, number>();
  forEachIdChunk(wordIds, (chunk) => {
    const rows = db
      .prepare(
        `
        SELECT o.word_id AS wordId, COUNT(DISTINCT sl.anime_id) AS animeCount
        FROM imm_word_line_occurrences o
        JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE o.word_id IN (${makePlaceholders(chunk)}) AND sl.anime_id IS NOT NULL
        GROUP BY o.word_id
      `,
      )
      .all(...chunk) as Array<{ wordId: number; animeCount: number }>;
    for (const row of rows) counts.set(row.wordId, row.animeCount);
  });
  return counts;
}

/**
 * New-word history is a SQL aggregate over first-seen dates rather than loading
 * every vocabulary row into the dashboard.
 */
export function getVocabularyChartData(db: DatabaseSync): VocabularyChartData {
  const ready = areLexicalDailyRollupsReady(db);
  const firstSeenRows = ready ? getNewWordFirstSeenRows(db) : [];
  const timeline = new Map<number, number>();
  const timelineWithoutNames = new Map<number, number>();
  for (const row of firstSeenRows) {
    timeline.set(row.firstSeenDay, (timeline.get(row.firstSeenDay) ?? 0) + row.wordCount);
    if (row.firstSeenDayWithoutNames !== null) {
      timelineWithoutNames.set(
        row.firstSeenDayWithoutNames,
        (timelineWithoutNames.get(row.firstSeenDayWithoutNames) ?? 0) + row.wordCount,
      );
    }
  }
  return {
    ready,
    newWordsTimeline: toNewWordTimeline(timeline),
    newWordsTimelineWithoutNames: toNewWordTimeline(timelineWithoutNames),
  };
}

function toNewWordTimeline(byDay: Map<number, number>): VocabularyChartData['newWordsTimeline'] {
  return [...byDay.entries()]
    .sort(([left], [right]) => left - right)
    .map(([epochDay, wordCount]) => ({ epochDay, wordCount }));
}

/** Word rows matching the user's stats exclusions, using the same aliases as the summary. */
export function getExcludedWordIds(db: DatabaseSync): number[] {
  const excludedAliases = new Set(
    getStatsExcludedWords(db).flatMap((word) => excludedVocabularyAliases(word)),
  );
  const aliases = [...excludedAliases];
  const ids = new Set<number>();
  for (let offset = 0; offset < aliases.length; offset += EXCLUSION_ALIAS_BATCH_SIZE) {
    const batch = aliases.slice(offset, offset + EXCLUSION_ALIAS_BATCH_SIZE);
    const placeholders = batch.map(() => '?').join(', ');
    const rows = db
      .prepare(
        `
          SELECT id AS wordId, headword, word, reading
          FROM imm_words
          WHERE vocabulary_visible = 1
            AND (headword IN (${placeholders}) OR word IN (${placeholders}) OR reading IN (${placeholders}))
        `,
      )
      .all(...batch, ...batch, ...batch) as Array<
      Pick<VocabularyStatsRow, 'headword' | 'word' | 'reading'> & { wordId: number }
    >;
    for (const row of rows) {
      if (excludedVocabularyAliases(row).some((alias) => excludedAliases.has(alias))) {
        ids.add(row.wordId);
      }
    }
  }
  return [...ids];
}

// The aggregate below scans every word (about 200ms on 68k rows), and the Trends tab
// runs it on each control change. Results stay valid until something writes:
// total_changes covers this connection, data_version covers other connections.
const newWordFirstSeenCache = new WeakMap<
  DatabaseSync,
  { version: string; rows: NewWordFirstSeenRow[] }
>();

export function databaseWriteVersion(db: DatabaseSync): string {
  const { changes } = db.prepare('SELECT total_changes() AS changes').get() as { changes: number };
  const { data_version: dataVersion } = db.prepare('PRAGMA data_version').get() as {
    data_version: number;
  };
  return `${changes}:${dataVersion}`;
}

/**
 * When each dictionary word was first seen, folded across its surface forms:
 * `imm_words` stores 言う and 言った as separate rows, so counting rows would count
 * conjugations as new words. Rows group identical (day, day without names, band)
 * results so the caller gets a small aggregate rather than one row per word.
 * `firstSeenDayWithoutNames` and `band` ignore proper-noun forms, and the day is
 * null when every form of the word is a proper noun.
 */
export function getNewWordFirstSeenRows(db: DatabaseSync): NewWordFirstSeenRow[] {
  const version = databaseWriteVersion(db);
  const cached = newWordFirstSeenCache.get(db);
  if (cached?.version === version) return cached.rows;

  const excludedIds = getExcludedWordIds(db);
  const exclusion =
    excludedIds.length > 0 ? `AND id NOT IN (${makePlaceholders(excludedIds)})` : '';
  const seconds = `CASE
      WHEN ABS(CAST(first_seen AS REAL)) >= 10000000000 THEN CAST(first_seen AS REAL) / 1000
      ELSE CAST(first_seen AS REAL)
    END`;
  const rows = db
    .prepare(
      `
        WITH dictionary_words AS (
          SELECT
            MIN(${seconds}) AS firstSeen,
            MIN(CASE WHEN ${NOT_A_NAME_SQL} THEN ${seconds} END) AS firstSeenWithoutNames,
            MIN(CASE WHEN ${NOT_A_NAME_SQL} THEN frequency_rank END) AS rank
          FROM imm_words
          WHERE vocabulary_visible = 1 AND first_seen IS NOT NULL ${exclusion}
          GROUP BY ${DICTIONARY_WORD_KEY_SQL}
        )
        SELECT
          ${localEpochDaySql('firstSeen')} AS firstSeenDay,
          CASE
            WHEN firstSeenWithoutNames IS NULL THEN NULL
            ELSE ${localEpochDaySql('firstSeenWithoutNames')}
          END AS firstSeenDayWithoutNames,
          CASE
            WHEN rank IS NULL THEN 'unranked'
            WHEN rank <= 1000 THEN 'top1k'
            WHEN rank <= 5000 THEN 'top5k'
            WHEN rank <= 10000 THEN 'top10k'
            WHEN rank <= 20000 THEN 'top20k'
            ELSE 'beyond20k'
          END AS band,
          COUNT(*) AS wordCount
        FROM dictionary_words
        GROUP BY 1, 2, 3
      `,
    )
    .all(...excludedIds) as unknown as NewWordFirstSeenRow[];
  newWordFirstSeenCache.set(db, { version, rows });
  return rows;
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
  // Exclusion and known-word rules live in JS, so rows are scanned in id-keyed
  // batches to keep memory bounded on large vocabularies. Once the lexical
  // rollups are ready, `vocabulary_visible` holds the persisted result of
  // isVocabularyStatsRowVisible (the same column the new-word charts count), so
  // the scan skips re-running the subtitle filter rules for every word.
  const persistedVisibility = areLexicalDailyRollupsReady(db);
  const scanStmt = db.prepare(`
    SELECT id AS wordId, headword, word, reading,
      part_of_speech AS partOfSpeech, pos1, pos2, pos3,
      frequency, frequency_rank AS frequencyRank,
      first_seen AS firstSeen, last_seen AS lastSeen,
      0 AS animeCount
    FROM imm_words
    WHERE id > ?${persistedVisibility ? ' AND vocabulary_visible = 1' : ''}
    ORDER BY id
    LIMIT ?
  `);
  const isVisible = persistedVisibility ? () => true : isVocabularyStatsRowVisible;
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
  // Keyed by dictionary form, so a word's conjugations count once and it is new
  // this week only if no form of it was seen earlier.
  const dictionaryWords = new Map<
    string,
    { firstSeenSec: number; firstSeenSecWithoutNames: number | null; known: boolean }
  >();

  let lastId = Number.MIN_SAFE_INTEGER;
  for (;;) {
    const words = scanStmt.all(lastId, scanBatchSize) as VocabularyStatsRow[];
    if (words.length === 0) break;
    lastId = words[words.length - 1]!.wordId;
    for (const word of words) {
      if (
        !isVisible(word) ||
        excludedVocabularyAliases(word).some((alias) => excludedAliases.has(alias))
      ) {
        continue;
      }
      const firstSeenSec = timestampSeconds(fromDbTimestamp(word.firstSeen) ?? 0);
      const key = vocabularyDisplayHeadword(word);
      const entry = dictionaryWords.get(key) ?? {
        firstSeenSec,
        firstSeenSecWithoutNames: null,
        known: false,
      };
      entry.known ||= (knownWords?.has(key) ?? false) || (knownWords?.has(word.word) ?? false);
      entry.firstSeenSec = Math.min(entry.firstSeenSec, firstSeenSec);
      if (!isProperNounPos2(word.pos2)) {
        entry.firstSeenSecWithoutNames = Math.min(
          entry.firstSeenSecWithoutNames ?? firstSeenSec,
          firstSeenSec,
        );
      }
      dictionaryWords.set(key, entry);
    }
    if (words.length < scanBatchSize) break;
  }

  for (const { firstSeenSec, firstSeenSecWithoutNames, known } of dictionaryWords.values()) {
    const hasNonNameForm = firstSeenSecWithoutNames !== null;
    summary.uniqueWords += 1;
    if (hasNonNameForm) summary.uniqueWordsWithoutNames += 1;
    if (firstSeenSec >= weekAgoSec) summary.newThisWeek += 1;
    if (hasNonNameForm && firstSeenSecWithoutNames >= weekAgoSec) {
      summary.newThisWeekWithoutNames += 1;
    }
    if (known) {
      summary.knownWordCount! += 1;
      if (hasNonNameForm) summary.knownWordCountWithoutNames! += 1;
    }
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

// Lines recorded before subtitle text was widened at ingest can still hold halfwidth
// katakana; widen them on the way out so the dashboard and cards mined from it match
// newer rows. Display-only: word identity columns stay as stored since they are keys.
function widenLineText<T extends { text: string }>(row: T): T {
  return { ...row, text: widenHalfwidthKana(row.text) };
}

export function getWordOccurrences(
  db: DatabaseSync,
  headword: string,
  word: string,
  reading: string,
  limit = 100,
  offset = 0,
  allForms = false,
): WordOccurrenceRow[] {
  const wordFilter = allForms
    ? `w.id IN (SELECT id FROM imm_words WHERE ${DICTIONARY_WORD_KEY_SQL} = ?)`
    : 'w.headword = ? AND w.word = ? AND w.reading = ?';
  const wordParams = allForms ? [headword] : [headword, word, reading];
  const rows = db
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
        WHERE ${wordFilter}
        ORDER BY ${MINEABLE_LINES_FIRST_SQL}, l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
        OFFSET ?
      `,
    )
    .all(...wordParams, limit, offset) as unknown as WordOccurrenceRow[];
  return rows.map(widenLineText);
}

export function getKanjiOccurrences(
  db: DatabaseSync,
  kanji: string,
  limit = 100,
  offset = 0,
): KanjiOccurrenceRow[] {
  const rows = db
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
  return rows.map(widenLineText);
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
    // Lines recorded before ingest widening can hold halfwidth katakana, so subtitle
    // text is matched against the term's widened and narrowed forms as well.
    const widenedTerm = widenHalfwidthKana(term);
    const textLikeTerms = uniqueNonEmptyTerms([
      term,
      widenedTerm,
      narrowFullwidthKana(widenedTerm),
    ]).map((variant) => `%${escapeLikeTerm(variant)}%`);
    const headwords = getHeadwordCandidatesForSentenceSearchTerm(term, options);
    // Uncorrelated on purpose: SQLite resolves the headword's lines once through
    // the word indexes. A per-line EXISTS probed every subtitle line (~0.2 s).
    const headwordClause =
      headwords.length > 0
        ? `
        OR l.line_id IN (
          SELECT o.line_id
          FROM imm_words w
          JOIN imm_word_line_occurrences o ON o.word_id = w.id
          WHERE w.headword IN (${headwords.map(() => '?').join(', ')})
        )
      `
        : '';
    clauses.push(`
      (
        ${textLikeTerms.map(() => "l.text LIKE ? ESCAPE '\\'").join(' OR ')}
        OR v.canonical_title LIKE ? ESCAPE '\\'
        OR COALESCE(a.canonical_title, '') LIKE ? ESCAPE '\\'
        ${headwordClause}
      )
    `);
    params.push(...textLikeTerms, likeTerm, likeTerm, ...headwords);
  }

  const rows = db
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
        ORDER BY ${MINEABLE_LINES_FIRST_SQL}, l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
      `,
    )
    .all(...params, resolvedLimit) as unknown as SentenceSearchResultRow[];
  return rows.map(widenLineText);
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

const WORD_DETAIL_FORM_LIMIT = 30;

/**
 * One entry per written form. A form can be stored more than once when its reading
 * was recorded differently, as 見て is under both みて and み. Returns most seen first.
 */
export function mergeFormsByWord<T extends { word: string; seen: number }>(forms: T[]): T[] {
  const byWord = new Map<string, T>();
  for (const form of forms) {
    const existing = byWord.get(form.word);
    if (existing) {
      existing.seen += form.seen;
    } else {
      byWord.set(form.word, { ...form });
    }
  }
  return [...byWord.values()].sort((left, right) => right.seen - left.seen);
}

/** A word's form plus totals across every form of its dictionary word. */
export function getWordDetail(db: DatabaseSync, wordId: number): WordDetailRow | null {
  const word = db
    .prepare(
      `
    SELECT id AS wordId, COALESCE(headword, '') AS headword, COALESCE(word, '') AS word,
           COALESCE(reading, '') AS reading,
           part_of_speech AS partOfSpeech, pos1, pos2, pos3
    FROM imm_words WHERE id = ?
  `,
    )
    .get(wordId) as Pick<
    WordDetailRow,
    'wordId' | 'headword' | 'word' | 'reading' | 'partOfSpeech' | 'pos1' | 'pos2' | 'pos3'
  > | null;
  if (!word) return null;

  const hasHeadword = word.headword.trim() !== '';
  const forms = db
    .prepare(
      `
    SELECT id AS wordId, COALESCE(word, '') AS word, COALESCE(reading, '') AS reading,
           frequency AS seen, frequency_rank AS rank,
           first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_words
    WHERE ${hasHeadword ? '(headword = ? AND vocabulary_visible = 1) OR id = ?' : 'id = ?'}
    ORDER BY frequency DESC, id
  `,
    )
    .all(...(hasHeadword ? [word.headword, wordId] : [wordId])) as Array<{
    wordId: number;
    word: string;
    reading: string;
    seen: number | null;
    rank: number | null;
    firstSeen: number;
    lastSeen: number;
  }>;

  const lines = db
    .prepare(
      `
    SELECT COUNT(DISTINCT l.line_id) AS lineCount,
           COUNT(DISTINCT CASE WHEN ${NO_LOCAL_FILE_SQL} THEN l.line_id END) AS streamLineCount
    FROM imm_word_line_occurrences o
    JOIN imm_words w ON w.id = o.word_id
    JOIN imm_subtitle_lines l ON l.line_id = o.line_id
    JOIN imm_videos v ON v.video_id = l.video_id
    WHERE ${hasHeadword ? 'w.headword = ?' : 'w.id = ?'}
  `,
    )
    .get(hasHeadword ? word.headword : wordId) as { lineCount: number; streamLineCount: number };
  const merged = mergeFormsByWord(forms.map((form) => ({ ...form, seen: form.seen ?? 0 })));
  const toSeconds = (value: number): number => timestampSeconds(fromDbTimestamp(value) ?? 0);
  const ranks = forms.map((form) => form.rank).filter((rank): rank is number => (rank ?? 0) > 0);
  return {
    ...word,
    frequency: forms.reduce((sum, form) => sum + (form.seen ?? 0), 0),
    frequencyRank: ranks.length > 0 ? Math.min(...ranks) : null,
    dictionaryReading:
      forms.find((form) => form.word === word.headword && form.reading.trim() !== '')?.reading ??
      '',
    formCount: merged.length,
    forms: merged.slice(0, WORD_DETAIL_FORM_LIMIT).map((form) => ({
      wordId: form.wordId,
      word: form.word,
      reading: form.reading,
      seen: form.seen,
    })),
    lineCount: lines.lineCount,
    streamLineCount: lines.streamLineCount,
    firstSeen: Math.min(...forms.map((form) => toSeconds(form.firstSeen))),
    lastSeen: Math.max(...forms.map((form) => toSeconds(form.lastSeen))),
  };
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
    JOIN imm_words w ON w.id = o.word_id
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE (w.id = ? OR w.headword = (SELECT NULLIF(TRIM(headword), '') FROM imm_words WHERE id = ?))
      AND sl.anime_id IS NOT NULL
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `,
    )
    .all(wordId, wordId) as WordAnimeAppearanceRow[];
}

/** Other dictionary words sharing a reading or kanji; a word's own forms don't count. */
export function getSimilarWords(db: DatabaseSync, wordId: number, limit = 10): SimilarWordRow[] {
  const word = db
    .prepare(
      `SELECT COALESCE(headword, '') AS headword, COALESCE(reading, '') AS reading
       FROM imm_words WHERE id = ?`,
    )
    .get(wordId) as { headword: string; reading: string } | null;
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

  const readingMatch =
    reading !== ''
      ? 'MAX(CASE WHEN reading = ? THEN 1 ELSE 0 END) OVER (PARTITION BY headword)'
      : '0';
  const readingMatchParams = reading !== '' ? [word.reading] : [];

  // One form matching is enough to relate a word, but its total and the form it shows
  // come from all its visible forms, so a rare form with the same reading can't stand in.
  // The blank check lists its characters because SQL TRIM alone strips only ASCII spaces.
  return db
    .prepare(
      `
    WITH related AS (
      SELECT DISTINCT headword
      FROM imm_words
      WHERE headword != ? AND vocabulary_visible = 1
      AND TRIM(headword, ' ' || char(9, 10, 11, 12, 13, 12288)) != ''
      AND (${clauses.join(' OR ')})
    ),
    forms AS (
      SELECT id, headword, word, reading,
             SUM(frequency) OVER (PARTITION BY headword) AS totalSeen,
             ${readingMatch} AS readingMatch,
             ROW_NUMBER() OVER (PARTITION BY headword ORDER BY frequency DESC, id) AS formOrder
      FROM imm_words
      WHERE vocabulary_visible = 1 AND headword IN (SELECT headword FROM related)
    )
    SELECT id AS wordId, headword, word, reading, totalSeen AS frequency
    FROM forms
    WHERE formOrder = 1
    ORDER BY readingMatch DESC, totalSeen DESC, headword
    LIMIT ?
  `,
    )
    .all(word.headword, ...params, ...readingMatchParams, limit) as SimilarWordRow[];
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
