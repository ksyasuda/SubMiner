import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import type { StatsDictionaryWordQuery } from '../../../types/stats-http-contract';
import {
  getSimilarWords,
  getWordAnimeAppearances,
  getVocabularySummary,
  getWordDetail,
  getWordOccurrences,
  replaceStatsExcludedWords,
  searchSubtitleSentences,
} from './query-lexical';
import { markLexicalDailyRollupsReady } from './lexical-rollups';
import { getDictionaryWordPage, getVocabularyCoverage } from './query-vocabulary-words';
import { startSessionRecord } from './session';
import { Database, type DatabaseSync } from './sqlite';
import { ensureSchema, getOrCreateAnimeRecord, getOrCreateVideoRecord } from './storage';
import { SOURCE_TYPE_LOCAL, SOURCE_TYPE_REMOTE } from './types';

function withDb(run: (db: DatabaseSync) => void): void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-vocabulary-'));
  const db = new Database(path.join(dir, 'immersion.sqlite'));
  try {
    ensureSchema(db);
    run(db);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

interface WordFixture {
  headword: string;
  word?: string;
  reading?: string;
  seen?: number;
  rank?: number | null;
  pos2?: string;
  partOfSpeech?: string;
  firstSeen?: number;
}

function addWord(db: DatabaseSync, word: WordFixture): number {
  const firstSeen = word.firstSeen ?? 1_700_000_000;
  const result = db
    .prepare(
      `INSERT INTO imm_words (
         headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen,
         frequency, frequency_rank
       ) VALUES (?, ?, ?, ?, '', ?, '', ?, ?, ?, ?)`,
    )
    .run(
      word.headword,
      word.word ?? word.headword,
      word.reading ?? '',
      word.partOfSpeech ?? 'noun',
      word.pos2 ?? '',
      firstSeen,
      firstSeen,
      word.seen ?? 1,
      word.rank ?? null,
    );
  return Number(result.lastInsertRowid);
}

function getIdOf(db: DatabaseSync, headword: string): number {
  const row = db.prepare('SELECT id FROM imm_words WHERE headword = ?').get(headword) as {
    id: number;
  };
  return row.id;
}

function addOccurrence(
  db: DatabaseSync,
  wordId: number,
  animeTitle: string,
  count = 1,
  source: 'file' | 'stream' = 'file',
  timing: { startMs: number | null; endMs: number | null } = { startMs: 1_000, endMs: 2_000 },
): void {
  const animeId = getOrCreateAnimeRecord(db, {
    parsedTitle: animeTitle,
    canonicalTitle: animeTitle,
    anilistId: null,
    titleRomaji: null,
    titleEnglish: null,
    titleNative: null,
    metadataJson: null,
  });
  const videoId = getOrCreateVideoRecord(
    db,
    source === 'file' ? `local:/tmp/${animeTitle}.mkv` : `remote:jellyfin://${animeTitle}`,
    {
      canonicalTitle: animeTitle,
      sourcePath: source === 'file' ? `/tmp/${animeTitle}.mkv` : null,
      sourceUrl: source === 'file' ? null : `jellyfin://${animeTitle}`,
      sourceType: source === 'file' ? SOURCE_TYPE_LOCAL : SOURCE_TYPE_REMOTE,
    },
  );
  const { sessionId } = startSessionRecord(db, videoId, 1_700_000_000_000);
  const lineId = Number(
    db
      .prepare(
        `INSERT INTO imm_subtitle_lines (
           session_id, video_id, anime_id, line_index, segment_start_ms, segment_end_ms, text
         ) VALUES (?, ?, ?, 0, ?, ?, ?)`,
      )
      .run(sessionId, videoId, animeId, timing.startMs, timing.endMs, `line ${animeTitle}`)
      .lastInsertRowid,
  );
  db.prepare(
    `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count) VALUES (?, ?, ?)`,
  ).run(lineId, wordId, count);
}

function query(overrides: Partial<StatsDictionaryWordQuery> = {}): StatsDictionaryWordQuery {
  return {
    sort: 'seen',
    descending: true,
    hideKnown: false,
    hideNames: false,
    hideKana: false,
    limit: 25,
    offset: 0,
    ...overrides,
  };
}

const headwords = (db: DatabaseSync, overrides: Partial<StatsDictionaryWordQuery>) =>
  getDictionaryWordPage(db, null, query(overrides)).rows.map((row) => row.headword);

test('conjugated forms fold into one word with its total and dictionary reading', () => {
  withDb((db) => {
    addWord(db, { headword: '見る', word: '見て', reading: 'みて', seen: 263, rank: 36 });
    const dictionaryId = addWord(db, {
      headword: '見る',
      word: '見る',
      reading: 'みる',
      seen: 172,
      rank: 36,
    });
    addWord(db, { headword: '見る', word: '見てる', reading: 'みてる', seen: 78, rank: 36 });

    const { rows, total } = getDictionaryWordPage(db, null, query());

    assert.equal(total, 1);
    assert.equal(rows[0]?.seen, 513);
    assert.equal(rows[0]?.reading, 'みる');
    assert.equal(rows[0]?.wordId, dictionaryId);
    assert.equal(rows[0]?.formCount, 3);
    assert.deepEqual(
      rows[0]?.forms.map((form) => form.word),
      ['見て', '見る', '見てる'],
    );
  });
});

test('a form stored under two readings counts as one form', () => {
  withDb((db) => {
    const first = addWord(db, { headword: '見る', word: '見て', reading: 'みて', seen: 263 });
    addWord(db, { headword: '見る', word: '見て', reading: 'み', seen: 122 });
    addWord(db, { headword: '見る', word: '見た', reading: 'みた', seen: 200 });

    const row = getDictionaryWordPage(db, null, query()).rows[0];
    assert.equal(row?.formCount, 2);
    assert.deepEqual(
      row?.forms.map((form) => [form.word, form.seen]),
      [
        ['見て', 385],
        ['見た', 200],
      ],
    );
    assert.equal(getWordDetail(db, first)?.formCount, 2);
  });
});

test('a word seen only conjugated has no reading rather than a conjugated one', () => {
  withDb((db) => {
    addWord(db, { headword: '待つ', word: '待って', reading: 'まって', seen: 5 });
    assert.equal(getDictionaryWordPage(db, null, query()).rows[0]?.reading, '');
  });
});

test('the list reaches rarely seen words, not just the most seen ones', () => {
  withDb((db) => {
    for (let index = 0; index < 600; index += 1) {
      addWord(db, { headword: `語${index}`, seen: 1_000 - index });
    }
    const { rows, total } = getDictionaryWordPage(db, null, query({ limit: 10, offset: 590 }));
    assert.equal(total, 600);
    assert.equal(rows[9]?.headword, '語599');
  });
});

test('search matches any form or reading, treating katakana and hiragana alike', () => {
  withDb((db) => {
    addWord(db, { headword: '猫', word: '猫', reading: 'ねこ' });
    addWord(db, { headword: '行く', word: '行った', reading: 'いった' });
    addWord(db, { headword: '犬', word: '犬', reading: 'いぬ' });

    assert.deepEqual(headwords(db, { search: 'ネコ' }), ['猫']);
    assert.deepEqual(headwords(db, { search: '行った' }), ['行く']);
    assert.deepEqual(headwords(db, { search: 'いった' }), ['行く']);
  });
});

for (const c of [
  { name: 'known words', filter: { hideKnown: true }, expected: ['海', '田中', 'なる'] },
  { name: 'names', filter: { hideNames: true }, expected: ['山', '海', 'なる'] },
  { name: 'kana', filter: { hideKana: true }, expected: ['山', '海', '田中'] },
  { name: 'a rank limit', filter: { maxRank: 500 }, expected: ['山', 'なる'] },
  { name: 'a part of speech', filter: { partOfSpeech: 'verb' as const }, expected: ['なる'] },
]) {
  test(`the list filters by ${c.name}`, () => {
    withDb((db) => {
      addWord(db, { headword: '山', seen: 40, rank: 100 });
      addWord(db, { headword: '海', seen: 30, rank: 9_000 });
      addWord(db, { headword: '田中', seen: 20, rank: 800, pos2: '固有名詞|一般' });
      addWord(db, { headword: 'なる', seen: 10, rank: 16, partOfSpeech: 'verb' });

      const rows = getDictionaryWordPage(db, new Set(['山']), query(c.filter)).rows;
      assert.deepEqual(
        rows.map((row) => row.headword),
        c.expected,
      );
    });
  });
}

test('a word with a non-name form is not a name and ranks as that form', () => {
  withDb((db) => {
    addWord(db, { headword: '白', word: 'しろ', rank: 49_000, pos2: '固有名詞' });
    addWord(db, { headword: '白', word: '白', rank: 1_500, pos2: '一般' });

    const row = getDictionaryWordPage(db, null, query({ hideNames: true })).rows[0];
    assert.equal(row?.isName, false);
    assert.equal(row?.rank, 1_500);
  });
});

test('excluded words stay out of the list', () => {
  withDb((db) => {
    addWord(db, { headword: '音', seen: 50 });
    addWord(db, { headword: '海', seen: 10 });
    assert.deepEqual(headwords(db, {}), ['音', '海']);

    replaceStatsExcludedWords(db, [{ headword: '音', word: '音', reading: '' }]);
    assert.deepEqual(headwords(db, {}), ['海']);
  });
});

test('unranked words sort last in both rank directions', () => {
  withDb((db) => {
    addWord(db, { headword: '無', rank: null });
    addWord(db, { headword: '高', rank: 10 });
    addWord(db, { headword: '低', rank: 900 });

    assert.deepEqual(headwords(db, { sort: 'rank', descending: false }), ['高', '低', '無']);
    assert.deepEqual(headwords(db, { sort: 'rank', descending: true }), ['低', '高', '無']);
  });
});

test('titles count distinct titles across every form and can filter and sort', () => {
  withDb((db) => {
    const seen = addWord(db, { headword: '見る', word: '見て', seen: 5 });
    const seenAgain = addWord(db, { headword: '見る', word: '見た', seen: 5 });
    const single = addWord(db, { headword: '海', seen: 50 });
    addOccurrence(db, seen, 'A');
    addOccurrence(db, seenAgain, 'B');
    addOccurrence(db, single, 'C');

    const page = getDictionaryWordPage(db, null, query());
    assert.deepEqual(
      page.rows.map((row) => [row.headword, row.titles]),
      [
        ['海', 1],
        ['見る', 2],
      ],
    );
    const across = getDictionaryWordPage(db, null, query({ sort: 'titles', minTitles: 2 }));
    assert.deepEqual(
      across.rows.map((row) => [row.headword, row.titles]),
      [['見る', 2]],
    );
  });
});

test('a word is stream only when every line, across its forms, came from a stream', () => {
  withDb((db) => {
    const streamed = addWord(db, { headword: '流', word: '流れ', seen: 30 });
    addOccurrence(db, streamed, 'Stream A', 1, 'stream');
    const mixed = addWord(db, { headword: '混', word: '混ぜ', seen: 20 });
    const mixedFile = addWord(db, { headword: '混', word: '混ぜた', seen: 10 });
    addOccurrence(db, mixed, 'Stream B', 1, 'stream');
    addOccurrence(db, mixedFile, 'File A', 1, 'file');
    addWord(db, { headword: '無', seen: 1 });

    const streamOnly = Object.fromEntries(
      getDictionaryWordPage(db, null, query()).rows.map((row) => [row.headword, row.streamOnly]),
    );
    // A word with no stored lines isn't blocked by streams.
    assert.deepEqual(streamOnly, { 流: true, 混: false, 無: false });
  });
});

test('word detail counts lines from streams across every form', () => {
  withDb((db) => {
    const streamed = addWord(db, { headword: '混', word: '混ぜ' });
    const file = addWord(db, { headword: '混', word: '混ぜた' });
    addOccurrence(db, streamed, 'Stream A', 3, 'stream');
    addOccurrence(db, streamed, 'Stream B', 1, 'stream');
    addOccurrence(db, file, 'File A', 2, 'file');

    const detail = getWordDetail(db, file);
    assert.equal(detail?.lineCount, 3);
    assert.equal(detail?.streamLineCount, 2);
  });
});

test('lines that can be mined come first in examples and search, newest first within each', () => {
  withDb((db) => {
    const wordId = addWord(db, { headword: '言う', word: '言った', reading: 'いった' });
    addOccurrence(db, wordId, 'File old', 1, 'file');
    addOccurrence(db, wordId, 'File new', 1, 'file');
    addOccurrence(db, wordId, 'File untimed', 1, 'file', { startMs: null, endMs: null });
    addOccurrence(db, wordId, 'Stream newest', 1, 'stream');

    const expected = ['File new', 'File old', 'Stream newest', 'File untimed'];
    assert.deepEqual(
      getWordOccurrences(db, '言う', '言った', 'いった', 10, 0, true).map((row) => row.videoTitle),
      expected,
    );
    assert.deepEqual(
      searchSubtitleSentences(db, 'line').map((row) => row.videoTitle),
      expected,
    );
  });
});

test('legacy rows with NULL text columns list, search, and open without failing', () => {
  withDb((db) => {
    const insert = db.prepare(
      `INSERT INTO imm_words (headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, 1700000000, 1700000000, 3)`,
    );
    const noHeadword = Number(insert.run(null, '猫', null).lastInsertRowid);
    const noForm = Number(insert.run('犬', null, null).lastInsertRowid);

    const { rows } = getDictionaryWordPage(db, null, query());
    assert.deepEqual(rows.map((row) => [row.headword, row.forms[0]?.word, row.reading]).sort(), [
      ['犬', '犬', ''],
      ['猫', '猫', ''],
    ]);
    // NULL must not become searchable text.
    assert.deepEqual(headwords(db, { search: 'null' }), []);

    for (const id of [noHeadword, noForm]) {
      assert.equal(getWordDetail(db, id)?.frequency, 3);
      assert.deepEqual(getWordAnimeAppearances(db, id), []);
      assert.deepEqual(getSimilarWords(db, id), []);
    }
  });
});

test('maxSeen and total describe every page, for scaling bars and paging', () => {
  withDb((db) => {
    addWord(db, { headword: '一', seen: 100 });
    addWord(db, { headword: '二', seen: 5 });

    const page = getDictionaryWordPage(db, null, query({ limit: 1, offset: 1 }));
    assert.equal(page.total, 2);
    assert.equal(page.maxSeen, 100);
    assert.deepEqual(
      page.rows.map((row) => row.headword),
      ['二'],
    );
  });
});

test('the summary counts a word as known when any of its forms is known', () => {
  withDb((db) => {
    markLexicalDailyRollupsReady(db);
    addWord(db, { headword: '見る', word: '見て', reading: 'みて' });
    addWord(db, { headword: '海' });

    assert.equal(getVocabularySummary(db, new Set(['見て'])).knownWordCount, 1);
    assert.equal(
      getDictionaryWordPage(db, new Set(['見て']), query()).rows.filter((row) => row.known).length,
      1,
    );
  });
});

test('coverage counts seen and known words per frequency band, without names', () => {
  withDb((db) => {
    addWord(db, { headword: '一', rank: 500 });
    addWord(db, { headword: '二', rank: 900 });
    addWord(db, { headword: '三', rank: 4_000 });
    addWord(db, { headword: '四', rank: 30_000 });
    addWord(db, { headword: '田中', rank: 700, pos2: '固有名詞' });
    addWord(db, { headword: '無', rank: null });

    assert.deepEqual(getVocabularyCoverage(db, new Set(['一', '三', '田中'])), [
      { band: 'top1k', seen: 2, known: 1 },
      { band: 'top5k', seen: 1, known: 1 },
      { band: 'top10k', seen: 0, known: 0 },
      { band: 'top20k', seen: 0, known: 0 },
    ]);
  });
});

test('word detail reports totals and a reading for the dictionary word, not one form', () => {
  withDb((db) => {
    const form = addWord(db, {
      headword: '見る',
      word: '見て',
      reading: 'みて',
      seen: 263,
      rank: 36,
      firstSeen: 1_700_000_500,
    });
    addWord(db, {
      headword: '見る',
      word: '見る',
      reading: 'みる',
      seen: 172,
      rank: 36,
      firstSeen: 1_700_000_100,
    });

    const detail = getWordDetail(db, form);

    assert.equal(detail?.word, '見て');
    assert.equal(detail?.frequency, 435);
    assert.equal(detail?.dictionaryReading, 'みる');
    assert.equal(detail?.frequencyRank, 36);
    assert.equal(detail?.formCount, 2);
    assert.equal(detail?.firstSeen, 1_700_000_100);
    assert.equal(detail?.lastSeen, 1_700_000_500);
  });
});

test('word appearances and examples can cover every form', () => {
  withDb((db) => {
    const first = addWord(db, { headword: '見る', word: '見て', reading: 'みて' });
    const second = addWord(db, { headword: '見る', word: '見た', reading: 'みた' });
    addOccurrence(db, first, 'A', 3);
    addOccurrence(db, second, 'B', 2);

    assert.deepEqual(
      getWordAnimeAppearances(db, first).map((row) => [row.animeTitle, row.occurrenceCount]),
      [
        ['A', 3],
        ['B', 2],
      ],
    );
  });
});

test("related words leave out a word's own forms and show one entry per word", () => {
  withDb((db) => {
    const form = addWord(db, { headword: '見る', word: '見て', reading: 'みて', seen: 10 });
    addWord(db, { headword: '見る', word: '見た', reading: 'みた', seen: 9 });
    addWord(db, { headword: '見える', word: '見えた', reading: 'みえた', seen: 4 });
    addWord(db, { headword: '見える', word: '見える', reading: 'みえる', seen: 3 });

    const similar = getSimilarWords(db, form);

    assert.deepEqual(
      similar.map((row) => [row.headword, row.frequency]),
      [['見える', 7]],
    );
    assert.equal(similar[0]?.word, '見えた');
  });
});

test('a reading match ranks a related word without choosing the form it shows', () => {
  withDb((db) => {
    const form = addWord(db, { headword: '見る', word: '見て', reading: 'みて', seen: 10 });
    // Matches on the kanji with every form, and on the reading with its minor form.
    const kanjiAndReading = addWord(db, {
      headword: '見える',
      word: '見えた',
      reading: 'みえた',
      seen: 9,
    });
    addWord(db, { headword: '見える', word: '見えて', reading: 'みて', seen: 2 });
    // Matches only on the reading of its minor form, so its other forms must still count.
    const readingOnly = addWord(db, {
      headword: '満ちる',
      word: '満ちた',
      reading: 'みちた',
      seen: 7,
    });
    addWord(db, { headword: '満ちる', word: '満ちて', reading: 'みて', seen: 1 });
    // More seen, but no reading match, so it still ranks after both.
    addWord(db, { headword: '見物', word: '見物', reading: 'けんぶつ', seen: 100 });
    // A hidden form supplies neither the shown form nor the total.
    const hidden = addWord(db, { headword: '満ちる', word: '満ち', reading: 'みち', seen: 50 });
    db.prepare('UPDATE imm_words SET vocabulary_visible = 0 WHERE id = ?').run(hidden);

    assert.deepEqual(
      getSimilarWords(db, form).map((row) => [row.wordId, row.word, row.reading, row.frequency]),
      [
        [kanjiAndReading, '見えた', 'みえた', 11],
        [readingOnly, '満ちた', 'みちた', 8],
        [getIdOf(db, '見物'), '見物', 'けんぶつ', 100],
      ],
    );
  });
});

test('related words show the same form when a word has tied top forms', () => {
  withDb((db) => {
    const form = addWord(db, { headword: '見る', word: '見て', reading: 'みて', seen: 10 });
    const first = addWord(db, { headword: '見える', word: '見えた', reading: 'みえた', seen: 5 });
    addWord(db, { headword: '見える', word: '見えて', reading: 'みえて', seen: 5 });

    assert.equal(getSimilarWords(db, form)[0]?.wordId, first);
  });
});

test('related words leave out rows with a blank headword', () => {
  withDb((db) => {
    const form = addWord(db, { headword: '見る', word: '見て', reading: 'みて', seen: 10 });
    for (const [index, headword] of ['', '  ', '\t', '\v', '\f', '　'].entries()) {
      addWord(db, { headword, word: `空${index}`, reading: 'みて', seen: 50 });
    }
    addWord(db, { headword: '満ちる', word: '満ちて', reading: 'みて', seen: 1 });

    assert.deepEqual(
      getSimilarWords(db, form).map((row) => row.headword),
      ['満ちる'],
    );
  });
});

test('all-forms examples include legacy rows stored without a headword', () => {
  withDb((db) => {
    const insert = db.prepare(
      `INSERT INTO imm_words (headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, 1700000000, 1700000000, 1)`,
    );
    const titlesFor = (headword: string): string[] =>
      getWordOccurrences(db, headword, '', '', 10, 0, true)
        .map((row) => row.videoTitle)
        .sort();

    addOccurrence(db, Number(insert.run('猫', '猫', 'ねこ').lastInsertRowid), 'Headword');
    addOccurrence(db, Number(insert.run(null, '猫', 'ねこ').lastInsertRowid), 'Null headword');
    addOccurrence(db, Number(insert.run('  ', '猫', 'ねこ').lastInsertRowid), 'Blank headword');
    addOccurrence(db, Number(insert.run(null, null, 'いぬ').lastInsertRowid), 'Reading only');
    addOccurrence(db, Number(insert.run('犬', '猫', 'いぬ').lastInsertRowid), 'Other headword');

    assert.deepEqual(titlesFor('猫'), ['Blank headword', 'Headword', 'Null headword']);
    assert.deepEqual(titlesFor('いぬ'), ['Reading only']);
    assert.deepEqual(titlesFor('犬'), ['Other headword']);
  });
});
