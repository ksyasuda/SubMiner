import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { getLexicalDailyRollups } from './lexical-rollups';
import { getTrendsDashboard } from './query-trends';
import { getVocabularyChartData } from './query-lexical';
import { Database } from './sqlite';
import { ensureSchema } from './storage';

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

test('vocabulary charts use complete top-word and lexical rollup data', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const insertWord = db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, '', 1700000000, 1700000000, ?)`,
    );
    for (let index = 0; index < 501; index += 1) {
      insertWord.run(`語${index}`, `語${index}`, index === 500 ? 10_000 : 1);
    }

    const charts = getVocabularyChartData(db);

    assert.equal(charts.topWords[0]?.headword, '語500');
    assert.equal(charts.topWords[0]?.frequency, 10_000);
    assert.equal(charts.newWordsTimeline[0]?.wordCount, 501);
  } finally {
    db.close();
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  }
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
