import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  areLexicalDailyRollupsReady,
  getLexicalDailyRollups,
  rebuildLexicalDailyRollups,
} from './lexical-rollups';
import { getTrendsDashboard } from './query-trends';
import {
  getVocabularyChartData,
  getVocabularySummary,
  replaceStatsExcludedWords,
} from './query-lexical';
import { Database } from './sqlite';
import type { DatabaseSync } from './sqlite';
import { createTrackerPreparedStatements, ensureSchema, executeQueuedWrite } from './storage';
import { BASE_MS, createDb, seedEndedSession, seedVideo } from './__tests__/lifetime-test-fixtures';

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-lexical-rollups-'));
  return path.join(dir, 'immersion.sqlite');
}

test('lexical daily rollups follow first-seen corrections and deletions', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const firstDay = 19_500;
    const correctedDay = firstDay + 2;
    const firstSeen = firstDay * 86_400 + 43_200;
    const correctedSeen = correctedDay * 86_400 + 43_200;

    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, ?, ?, 1)`,
    ).run('猫', '猫', 'ねこ', firstSeen, firstSeen);
    db.prepare(
      `INSERT INTO imm_kanji(kanji, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, 1)`,
    ).run('猫', firstSeen, firstSeen);

    assert.deepEqual(getLexicalDailyRollups(db), [
      { epochDay: firstDay, wordCount: 1, wordCountWithoutNames: 1, kanjiCount: 1 },
    ]);

    db.prepare(`UPDATE imm_words SET first_seen = ? WHERE headword = ?`).run(correctedSeen, '猫');
    db.prepare(`DELETE FROM imm_kanji WHERE kanji = ?`).run('猫');

    assert.deepEqual(getLexicalDailyRollups(db), [
      { epochDay: correctedDay, wordCount: 1, wordCountWithoutNames: 1, kanjiCount: 0 },
    ]);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('lexical daily rollups normalize second and millisecond timestamps', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const epochDay = 19_500;
    const timestampSeconds = epochDay * 86_400 + 43_200;
    const timestampMilliseconds = timestampSeconds * 1_000;

    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, ?, ?, 1)`,
    ).run('猫', '猫', 'ねこ', timestampSeconds, timestampSeconds);
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, ?, ?, 1)`,
    ).run('犬', '犬', 'いぬ', timestampMilliseconds, timestampMilliseconds);
    db.prepare(
      `INSERT INTO imm_kanji(kanji, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, 1)`,
    ).run('猫', timestampSeconds, timestampSeconds);
    db.prepare(
      `INSERT INTO imm_kanji(kanji, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, 1)`,
    ).run('犬', timestampMilliseconds, timestampMilliseconds);

    assert.deepEqual(getLexicalDailyRollups(db), [
      { epochDay, wordCount: 2, wordCountWithoutNames: 2, kanjiCount: 2 },
    ]);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('lexical rollup rebuild excludes rows hidden by vocabulary persistence rules', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const epochDay = 19_500;
    const firstSeen = epochDay * 86_400 + 43_200;
    db.prepare(
      `INSERT INTO imm_words(
         headword, word, reading, part_of_speech, first_seen, last_seen, frequency
       ) VALUES (?, ?, ?, ?, ?, ?, 1)`,
    ).run('猫', '猫', 'ねこ', 'noun', firstSeen, firstSeen);
    db.prepare(
      `INSERT INTO imm_words(
         headword, word, reading, part_of_speech, first_seen, last_seen, frequency
       ) VALUES (?, ?, ?, ?, ?, ?, 1)`,
    ).run('は', 'は', 'は', 'particle', firstSeen, firstSeen);

    rebuildLexicalDailyRollups(db);

    assert.deepEqual(getLexicalDailyRollups(db), [
      { epochDay, wordCount: 1, wordCountWithoutNames: 1, kanjiCount: 0 },
    ]);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('re-seeing a filtered legacy word keeps charts consistent with the summary', () => {
  const db = createDb();

  try {
    const videoId = seedVideo(db, null, 'episode');
    const sessionId = seedEndedSession(db, videoId, BASE_MS, { activeMs: 1_000 });
    const firstSeen = BASE_MS / 1000;
    // Legacy row whose stored merged-token POS fails the persistence filter.
    db.prepare(
      `INSERT INTO imm_words(
         headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
       ) VALUES ('二死', 'にし', 'にし', 'other', '助詞|動詞', '格助詞|自立', '一般|*', ?, ?, 1)`,
    ).run(firstSeen, firstSeen);
    rebuildLexicalDailyRollups(db);

    // The incoming token passes the filter, but the upsert keeps the stored POS.
    executeQueuedWrite(
      {
        kind: 'subtitleLine',
        sessionId,
        videoId,
        lineIndex: 1,
        segmentStartMs: null,
        segmentEndMs: null,
        text: 'にし',
        wordOccurrences: [
          {
            headword: '二死',
            word: 'にし',
            reading: 'にし',
            partOfSpeech: 'verb',
            pos1: '動詞',
            pos2: '自立',
            pos3: '*',
            frequencyRank: null,
            occurrenceCount: 1,
          },
        ],
        kanjiOccurrences: [],
        firstSeen: firstSeen + 60,
        lastSeen: firstSeen + 60,
      },
      createTrackerPreparedStatements(db),
    );

    const plotted = getVocabularyChartData(db).newWordsTimeline.reduce(
      (sum, day) => sum + day.wordCount,
      0,
    );
    assert.equal(getVocabularySummary(db, null).uniqueWords, 0);
    assert.equal(plotted, 0);
  } finally {
    db.close();
  }
});

test('lexical rollup rebuild repairs last_seen left pointing at deleted history', () => {
  const db = createDb();

  try {
    const videoId = seedVideo(db, null, 'episode');
    const sessionId = seedEndedSession(db, videoId, BASE_MS, { activeMs: 1_000 });
    const lineId = Number(
      db
        .prepare(
          `INSERT INTO imm_subtitle_lines(session_id, video_id, line_index, text, CREATED_DATE, LAST_UPDATE_DATE)
           VALUES (?, ?, 1, '猫', ?, ?)`,
        )
        .run(sessionId, videoId, BASE_MS, BASE_MS).lastInsertRowid,
    );
    const seen = (BASE_MS + 251) / 1000;
    // first_seen this early may be real history pruned by retention; never touched.
    const earlyFirst = seen - 86_400.5;
    const staleLast = seen + 86_400.631;
    db.prepare(
      `INSERT INTO imm_words(id, headword, word, reading, first_seen, last_seen, frequency)
       VALUES (1, '猫', '猫', 'ねこ', ?, ?, 1), (2, '犬', '犬', 'いぬ', ?, ?, 1)`,
    ).run(earlyFirst, staleLast, seen - 0.1, seen - 0.1);
    db.prepare(
      `INSERT INTO imm_kanji(id, kanji, first_seen, last_seen, frequency) VALUES (1, '猫', ?, ?, 1)`,
    ).run(earlyFirst, staleLast);
    db.prepare(
      `INSERT INTO imm_word_line_occurrences(line_id, word_id, occurrence_count, seen_ms)
       VALUES (?, 1, 1, ?), (?, 2, 1, ?)`,
    ).run(lineId, BASE_MS + 251, lineId, BASE_MS + 251);
    db.prepare(
      `INSERT INTO imm_kanji_line_occurrences(line_id, kanji_id, occurrence_count, seen_ms)
       VALUES (?, 1, 1, ?)`,
    ).run(lineId, BASE_MS + 251);

    rebuildLexicalDailyRollups(db);

    const bounds = (table: string): Array<[number, number]> =>
      (
        db
          .prepare(`SELECT first_seen AS first, last_seen AS last FROM ${table} ORDER BY id`)
          .all() as Array<{
          first: number;
          last: number;
        }>
      ).map((row) => [row.first, row.last]);
    assert.deepEqual(
      bounds('imm_words'),
      [
        [earlyFirst, seen],
        [seen - 0.1, seen - 0.1],
      ],
      'healthy rows are untouched',
    );
    assert.deepEqual(bounds('imm_kanji'), [[earlyFirst, seen]]);
  } finally {
    db.close();
  }
});

test('lexical rollup rebuild tolerates nullable legacy vocabulary text', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (NULL, NULL, NULL, 1700000000, 1700000000, 1)`,
    ).run();
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (NULL, '猫', 'ねこ', 1700000000, 1700000000, 1)`,
    ).run();

    assert.doesNotThrow(() => rebuildLexicalDailyRollups(db));
    assert.equal(areLexicalDailyRollupsReady(db), true);
    assert.equal(getVocabularySummary(db, null).uniqueWords, 1);
    assert.equal(getVocabularySummary(db, new Set(['猫'])).knownWordCount, 1);
    assert.equal(getVocabularyChartData(db).topWords[0]?.headword, '猫');
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('lexical rollup rebuild scans vocabulary visibility in bounded id batches', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  const expectedBatchSize = 5_000;

  try {
    ensureSchema(db);
    const insertWord = db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, '', 1700000000, 1700000000, 1)`,
    );
    db.exec('BEGIN');
    for (let index = 0; index <= expectedBatchSize; index += 1) {
      insertWord.run(`語${index}`, `語${index}`);
    }
    db.exec('COMMIT');

    const scanPageSizes: number[] = [];
    const instrumentedDb: DatabaseSync = {
      prepare(source) {
        const statement = db.prepare(source);
        if (!source.includes('WHERE id > ?') || !source.includes('ORDER BY id')) {
          return statement;
        }
        return {
          run: (...params) => statement.run(...params),
          get: (...params) => statement.get(...params),
          all: (...params) => {
            const rows = statement.all(...params);
            scanPageSizes.push(rows.length);
            return rows;
          },
        };
      },
      exec(source) {
        db.exec(source);
        return instrumentedDb;
      },
      close() {
        return instrumentedDb;
      },
    };

    rebuildLexicalDailyRollups(instrumentedDb);

    assert.deepEqual(scanPageSizes, [expectedBatchSize, 1]);
    assert.equal(getVocabularySummary(db, null).uniqueWords, expectedBatchSize + 1);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('chart exclusions do not subtract vocabulary rows already hidden from the rollup', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const epochDay = 19_500;
    const firstSeen = epochDay * 86_400 + 43_200;
    db.prepare(
      `INSERT INTO imm_words(
         headword, word, reading, part_of_speech, first_seen, last_seen, frequency
       ) VALUES (?, ?, ?, ?, ?, ?, 1)`,
    ).run('猫', '猫', 'ねこ', 'noun', firstSeen, firstSeen);
    db.prepare(
      `INSERT INTO imm_words(
         headword, word, reading, part_of_speech, first_seen, last_seen, frequency
       ) VALUES (?, ?, ?, ?, ?, ?, 1)`,
    ).run('は', 'は', 'は', 'particle', firstSeen, firstSeen);
    rebuildLexicalDailyRollups(db);
    replaceStatsExcludedWords(db, [{ headword: 'は', word: 'は', reading: 'は' }]);

    const charts = getVocabularyChartData(db);

    assert.deepEqual(charts.newWordsTimeline, [{ epochDay, wordCount: 1 }]);
    assert.deepEqual(charts.newWordsTimelineWithoutNames, [{ epochDay, wordCount: 1 }]);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('legacy lexical rollup readiness does not satisfy the current rollup version', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_rollup_state(state_key, state_value)
       VALUES ('lexical_daily_rollups_ready', '1')
       ON CONFLICT(state_key) DO UPDATE SET state_value = excluded.state_value`,
    ).run();
    db.prepare(
      `DELETE FROM imm_rollup_state WHERE state_key = 'lexical_daily_rollups_version'`,
    ).run();

    assert.equal(areLexicalDailyRollupsReady(db), false);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('current lexical rollup readiness accepts legacy integer state storage', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    db.exec(`
      CREATE TABLE imm_rollup_state(
        state_key TEXT PRIMARY KEY,
        state_value INTEGER NOT NULL
      );
      INSERT INTO imm_rollup_state(state_key, state_value)
      VALUES ('lexical_daily_rollups_version', 3);
    `);

    assert.equal(areLexicalDailyRollupsReady(db), true);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('imm_words persists vocabulary visibility for rollup maintenance', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const columns = db.prepare(`PRAGMA table_info(imm_words)`).all() as Array<{ name: string }>;

    assert.equal(
      columns.some((column) => column.name === 'vocabulary_visible'),
      true,
    );
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('vocabulary charts use complete top-word and lexical rollup data', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const insertWord = db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, '', 1700000000, 1700000000, ?)`,
    );
    db.exec('BEGIN');
    for (let index = 0; index < 501; index += 1) {
      insertWord.run(`語${index}`, `語${index}`, index === 500 ? 10_000 : 1);
    }
    db.exec('COMMIT');

    const charts = getVocabularyChartData(db);

    assert.equal(charts.topWords[0]?.headword, '語500');
    assert.equal(charts.topWords[0]?.frequency, 10_000);
    assert.equal(charts.newWordsTimeline[0]?.wordCount, 501);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('vocabulary charts find full top-word sets beyond excluded and name rows', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const insertWord = db.prepare(
      `INSERT INTO imm_words(headword, word, reading, pos2, first_seen, last_seen, frequency)
       VALUES (?, ?, '', ?, 1700000000, 1700000000, ?)`,
    );
    const exclusions = [];
    for (let index = 0; index < 100; index += 1) {
      const headword = `語${index}`;
      insertWord.run(
        headword,
        headword,
        index < 80 && index >= 60 ? '固有名詞' : '一般',
        100 - index,
      );
      if (index < 60) exclusions.push({ headword, word: headword, reading: '' });
    }
    replaceStatsExcludedWords(db, exclusions);

    const charts = getVocabularyChartData(db);

    assert.equal(charts.topWords.length, 12);
    assert.equal(charts.topWords[0]?.headword, '語60');
    assert.equal(charts.topWordsWithoutNames.length, 12);
    assert.equal(charts.topWordsWithoutNames[0]?.headword, '語80');
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('vocabulary charts handle exclusion lists above one SQLite variable batch', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES ('語0', '語0', '', 1700000000, 1700000000, 1)`,
    ).run();
    const exclusions = Array.from({ length: 10_923 }, (_, index) => ({
      headword: `語${index}`,
      word: `語${index}`,
      reading: '',
    }));
    replaceStatsExcludedWords(db, exclusions);

    const charts = getVocabularyChartData(db);

    assert.deepEqual(charts.topWords, []);
    assert.deepEqual(charts.newWordsTimeline, []);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});

test('lexical rollup rebuild preserves the original error when rollback also fails', () => {
  const originalError = new Error('rebuild failed');
  const db = {
    exec(sql: string) {
      if (sql === 'BEGIN IMMEDIATE') return;
      if (sql === 'ROLLBACK') throw new Error('rollback failed');
      throw originalError;
    },
    prepare() {
      return { all: () => [], run: () => undefined };
    },
  } as unknown as DatabaseSync;

  assert.throws(() => rebuildLexicalDailyRollups(db), originalError);
});

test('trends read historical new-word buckets from lexical rollups', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES ('海', '海', 'うみ', 1700000000, 1700000000, 1)`,
    ).run();
    db.prepare(`UPDATE imm_lexical_daily_rollups SET word_count = 9`).run();

    const dashboard = getTrendsDashboard(db, 'all', 'day', false);

    assert.equal(dashboard.progress.newWords[0]?.value, 9);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
});
