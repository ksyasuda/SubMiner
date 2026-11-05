import type {
  StatsDictionaryWordPos,
  StatsDictionaryWordQuery,
} from '../../../types/stats-http-contract';
import type {
  DictionaryWordForm,
  DictionaryWordPage,
  DictionaryWordRow,
  VocabularyCoverageBand,
  VocabularyCoverageBandId,
} from '../../../types/stats-wire';
import type { DatabaseSync } from './sqlite';
import {
  databaseWriteVersion,
  DICTIONARY_WORD_KEY_SQL,
  getExcludedWordIds,
  isProperNounPos2,
  mergeFormsByWord,
  NO_LOCAL_FILE_SQL,
} from './query-lexical';
import { forEachIdChunk, makePlaceholders } from './query-shared';
import { nowMs } from './time';

const MAX_FORMS_PER_ROW = 8;
// Counting titles for every word walks every occurrence (about 900ms on 500k rows),
// and playback writes constantly, so that count tolerates being this stale.
const ALL_TITLE_COUNTS_MAX_AGE_MS = 5 * 60_000;

const COVERAGE_BAND_LIMITS: ReadonlyArray<{ band: VocabularyCoverageBandId; maxRank: number }> = [
  { band: 'top1k', maxRank: 1_000 },
  { band: 'top5k', maxRank: 5_000 },
  { band: 'top10k', maxRank: 10_000 },
  { band: 'top20k', maxRank: 20_000 },
];

const PART_OF_SPEECH_GROUPS: Record<StatsDictionaryWordPos, ReadonlySet<string>> = {
  noun: new Set(['noun']),
  verb: new Set(['verb']),
  adjective: new Set(['i_adjective', 'na_adjective']),
};

const KANA_ONLY_TEXT = /^[\p{Script=Hiragana}\p{Script=Katakana}ーゝゞヽヾ]+$/u;

interface DictionaryWordFormRow {
  id: number;
  key: string;
  word: string;
  reading: string;
  seen: number | null;
  rank: number | null;
  pos2: string | null;
  partOfSpeech: string | null;
  firstSeen: number | null;
  lastSeen: number | null;
}

interface DictionaryWord {
  key: string;
  seen: number;
  rank: number | null;
  isName: boolean;
  isKana: boolean;
  firstSeen: number;
  lastSeen: number;
  searchText: string;
  /** Most seen first. */
  forms: DictionaryWordFormRow[];
}

const wordCache = new WeakMap<DatabaseSync, { version: string; words: DictionaryWord[] }>();
const allTitleCountsCache = new WeakMap<
  DatabaseSync,
  { version: string; computedAtMs: number; counts: Map<string, number> }
>();

function toSeconds(value: number | null): number {
  const numeric = Number(value ?? 0);
  return Math.abs(numeric) >= 10_000_000_000 ? numeric / 1000 : numeric;
}

/** Lowercase hiragana, so a search matches regardless of kana type or case. */
function normalizeSearchText(text: string): string {
  return text
    .toLowerCase()
    .replace(/[ァ-ヶ]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0x60));
}

function loadDictionaryWords(db: DatabaseSync): DictionaryWord[] {
  const version = databaseWriteVersion(db);
  const cached = wordCache.get(db);
  if (cached?.version === version) return cached.words;

  const excludedIds = getExcludedWordIds(db);
  const exclusion =
    excludedIds.length > 0 ? `AND id NOT IN (${makePlaceholders(excludedIds)})` : '';
  const rows = db
    .prepare(
      `
        SELECT id, ${DICTIONARY_WORD_KEY_SQL} AS key,
          COALESCE(NULLIF(TRIM(word), ''), ${DICTIONARY_WORD_KEY_SQL}) AS word,
          COALESCE(reading, '') AS reading, frequency AS seen,
          frequency_rank AS rank, pos2, part_of_speech AS partOfSpeech,
          first_seen AS firstSeen, last_seen AS lastSeen
        FROM imm_words
        WHERE vocabulary_visible = 1 ${exclusion}
        ORDER BY frequency DESC, id
      `,
    )
    .all(...excludedIds) as unknown as DictionaryWordFormRow[];

  const byKey = new Map<string, DictionaryWord>();
  for (const row of rows) {
    if (!row.key) continue;
    const firstSeen = toSeconds(row.firstSeen);
    const lastSeen = toSeconds(row.lastSeen);
    const isName = isProperNounPos2(row.pos2);
    const entry = byKey.get(row.key);
    if (entry) {
      entry.seen += row.seen ?? 0;
      entry.forms.push(row);
      entry.firstSeen = Math.min(entry.firstSeen, firstSeen);
      entry.lastSeen = Math.max(entry.lastSeen, lastSeen);
      entry.isName &&= isName;
      entry.searchText += `\n${row.word}\n${row.reading}`;
    } else {
      byKey.set(row.key, {
        key: row.key,
        seen: row.seen ?? 0,
        rank: null,
        isName,
        isKana: KANA_ONLY_TEXT.test(row.key),
        firstSeen,
        lastSeen,
        searchText: `${row.key}\n${row.word}\n${row.reading}`,
        forms: [row],
      });
    }
  }

  const words = [...byKey.values()];
  for (const word of words) {
    word.searchText = normalizeSearchText(word.searchText);
    word.rank = bestRank(word.forms);
  }
  wordCache.set(db, { version, words });
  return words;
}

/** Rank from the word's non-name forms when it has any, so 白 ranks as the color. */
function bestRank(forms: DictionaryWordFormRow[]): number | null {
  const nonNames = forms.filter((form) => !isProperNounPos2(form.pos2));
  let best: number | null = null;
  for (const form of nonNames.length > 0 ? nonNames : forms) {
    if (form.rank !== null && form.rank > 0 && (best === null || form.rank < best)) {
      best = form.rank;
    }
  }
  return best;
}

function isKnown(word: DictionaryWord, knownWords: ReadonlySet<string> | null): boolean {
  if (!knownWords) return false;
  return knownWords.has(word.key) || word.forms.some((form) => knownWords.has(form.word));
}

function dictionaryForm(word: DictionaryWord): DictionaryWordFormRow | undefined {
  return word.forms.find((form) => form.word === word.key);
}

function getAllTitleCounts(db: DatabaseSync): Map<string, number> {
  const version = databaseWriteVersion(db);
  const cached = allTitleCountsCache.get(db);
  if (
    cached &&
    (cached.version === version || nowMs() - cached.computedAtMs < ALL_TITLE_COUNTS_MAX_AGE_MS)
  ) {
    return cached.counts;
  }
  const rows = db
    .prepare(
      `
        SELECT ${DICTIONARY_WORD_KEY_SQL} AS key, COUNT(DISTINCT sl.anime_id) AS titles
        FROM imm_words
        JOIN imm_word_line_occurrences o ON o.word_id = imm_words.id
        JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE sl.anime_id IS NOT NULL AND vocabulary_visible = 1
        GROUP BY key
      `,
    )
    .all() as Array<{ key: string; titles: number }>;
  const counts = new Map(rows.map((row) => [row.key, row.titles]));
  allTitleCountsCache.set(db, { version, computedAtMs: nowMs(), counts });
  return counts;
}

function getTitleCountsForWords(db: DatabaseSync, words: DictionaryWord[]): Map<string, number> {
  const keyByWordId = new Map<number, string>();
  for (const word of words) {
    for (const form of word.forms) keyByWordId.set(form.id, word.key);
  }
  const animeByKey = new Map<string, Set<number>>();
  forEachIdChunk([...keyByWordId.keys()], (chunk) => {
    const rows = db
      .prepare(
        `
          SELECT DISTINCT o.word_id AS wordId, sl.anime_id AS animeId
          FROM imm_word_line_occurrences o
          JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
          WHERE o.word_id IN (${makePlaceholders(chunk)}) AND sl.anime_id IS NOT NULL
        `,
      )
      .all(...chunk) as Array<{ wordId: number; animeId: number }>;
    for (const row of rows) {
      const key = keyByWordId.get(row.wordId);
      if (key === undefined) continue;
      const anime = animeByKey.get(key) ?? new Set<number>();
      anime.add(row.animeId);
      animeByKey.set(key, anime);
    }
  });
  return new Map([...animeByKey].map(([key, anime]) => [key, anime.size]));
}

/** Words whose every stored line came from a stream. Words with no lines don't count. */
function getStreamOnlyWords(db: DatabaseSync, words: DictionaryWord[]): Set<string> {
  const keyByWordId = new Map<number, string>();
  for (const word of words) {
    for (const form of word.forms) keyByWordId.set(form.id, word.key);
  }
  const withLines = new Set<string>();
  const withFile = new Set<string>();
  forEachIdChunk([...keyByWordId.keys()], (chunk) => {
    const rows = db
      .prepare(
        `
          SELECT o.word_id AS wordId, MIN(${NO_LOCAL_FILE_SQL}) AS allStreamed
          FROM imm_word_line_occurrences o
          JOIN imm_subtitle_lines l ON l.line_id = o.line_id
          JOIN imm_videos v ON v.video_id = l.video_id
          WHERE o.word_id IN (${makePlaceholders(chunk)})
          GROUP BY o.word_id
        `,
      )
      .all(...chunk) as Array<{ wordId: number; allStreamed: number }>;
    for (const row of rows) {
      const key = keyByWordId.get(row.wordId);
      if (key === undefined) continue;
      withLines.add(key);
      if (row.allStreamed === 0) withFile.add(key);
    }
  });
  return new Set([...withLines].filter((key) => !withFile.has(key)));
}

function compareWords(
  left: DictionaryWord,
  right: DictionaryWord,
  query: StatsDictionaryWordQuery,
  titleCounts: Map<string, number> | null,
): number {
  const sign = query.descending ? -1 : 1;
  const byText = left.key < right.key ? -1 : left.key > right.key ? 1 : 0;
  switch (query.sort) {
    case 'headword':
      return byText * sign;
    case 'rank':
      // Unranked words go last in either direction.
      if (left.rank === null || right.rank === null) {
        return (left.rank === null ? 1 : 0) - (right.rank === null ? 1 : 0) || byText;
      }
      return (left.rank - right.rank) * sign || byText;
    case 'titles':
      return (
        ((titleCounts?.get(left.key) ?? 0) - (titleCounts?.get(right.key) ?? 0)) * sign ||
        right.seen - left.seen ||
        byText
      );
    case 'firstSeen':
      return (left.firstSeen - right.firstSeen) * sign || byText;
    case 'lastSeen':
      return (left.lastSeen - right.lastSeen) * sign || byText;
    case 'seen':
      return (left.seen - right.seen) * sign || byText;
  }
}

function toRow(
  word: DictionaryWord,
  knownWords: ReadonlySet<string> | null,
  titleCounts: Map<string, number>,
  streamOnlyWords: ReadonlySet<string>,
): DictionaryWordRow {
  const dictionary = dictionaryForm(word);
  const topForm = word.forms[0]!;
  const reading = word.forms.find(
    (form) => form.word === word.key && form.reading.trim() !== '',
  )?.reading;
  const forms = mergeFormsByWord<DictionaryWordForm>(
    word.forms.map((form) => ({
      wordId: form.id,
      word: form.word,
      reading: form.reading,
      seen: form.seen ?? 0,
    })),
  );
  return {
    wordId: (dictionary ?? topForm).id,
    headword: word.key,
    reading: reading ?? '',
    partOfSpeech: topForm.partOfSpeech,
    rank: word.rank,
    seen: word.seen,
    titles: titleCounts.get(word.key) ?? 0,
    firstSeen: word.firstSeen,
    lastSeen: word.lastSeen,
    known: isKnown(word, knownWords),
    isName: word.isName,
    streamOnly: streamOnlyWords.has(word.key),
    formCount: forms.length,
    forms: forms.slice(0, MAX_FORMS_PER_ROW),
  };
}

/** One page of dictionary words, folded across their forms and filtered on the server. */
export function getDictionaryWordPage(
  db: DatabaseSync,
  knownWords: ReadonlySet<string> | null,
  query: StatsDictionaryWordQuery,
): DictionaryWordPage {
  const needle = query.search ? normalizeSearchText(query.search) : '';
  const posGroup = query.partOfSpeech ? PART_OF_SPEECH_GROUPS[query.partOfSpeech] : null;
  const needsAllTitleCounts = query.sort === 'titles' || query.minTitles !== undefined;
  const allTitleCounts = needsAllTitleCounts ? getAllTitleCounts(db) : null;

  const matches = loadDictionaryWords(db).filter((word) => {
    if (query.hideNames && word.isName) return false;
    if (query.hideKana && word.isKana) return false;
    if (query.maxRank !== undefined && (word.rank === null || word.rank > query.maxRank)) {
      return false;
    }
    if (posGroup && !posGroup.has(word.forms[0]?.partOfSpeech ?? '')) return false;
    if (needle && !word.searchText.includes(needle)) return false;
    if (query.minTitles !== undefined && (allTitleCounts?.get(word.key) ?? 0) < query.minTitles) {
      return false;
    }
    return !(query.hideKnown && isKnown(word, knownWords));
  });

  matches.sort((left, right) => compareWords(left, right, query, allTitleCounts));
  const pageWords = matches.slice(query.offset, query.offset + query.limit);
  const titleCounts = allTitleCounts ?? getTitleCountsForWords(db, pageWords);
  const streamOnlyWords = getStreamOnlyWords(db, pageWords);
  return {
    rows: pageWords.map((word) => toRow(word, knownWords, titleCounts, streamOnlyWords)),
    total: matches.length,
    maxSeen: matches.reduce((max, word) => Math.max(max, word.seen), 0),
  };
}

/** How many seen words in each frequency band are known. Names don't count. */
export function getVocabularyCoverage(
  db: DatabaseSync,
  knownWords: ReadonlySet<string>,
): VocabularyCoverageBand[] {
  const bands = COVERAGE_BAND_LIMITS.map(({ band }) => ({ band, seen: 0, known: 0 }));
  for (const word of loadDictionaryWords(db)) {
    if (word.isName || word.rank === null) continue;
    const index = COVERAGE_BAND_LIMITS.findIndex(({ maxRank }) => word.rank! <= maxRank);
    const band = bands[index];
    if (!band) continue;
    band.seen += 1;
    if (isKnown(word, knownWords)) band.known += 1;
  }
  return bands;
}
