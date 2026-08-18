import type { DatabaseSync } from './sqlite';
import { isVocabularyStatsRowVisible, type VocabularyVisibilityRow } from './vocabulary-visibility';

export interface LexicalDailyRollup {
  epochDay: number;
  wordCount: number;
  wordCountWithoutNames: number;
  kanjiCount: number;
}

const LOCAL_EPOCH_DAY_SQL = `
  CAST(julianday(
    CASE
      WHEN ABS(CAST(%VALUE% AS REAL)) >= 10000000000 THEN CAST(%VALUE% AS REAL) / 1000
      ELSE CAST(%VALUE% AS REAL)
    END,
    'unixepoch', 'localtime'
  ) - 2440587.5 AS INTEGER)
`;

const LEXICAL_DAILY_ROLLUP_VERSION = '2';
const LEXICAL_DAILY_ROLLUP_VERSION_KEY = 'lexical_daily_rollups_version';
const VOCABULARY_VISIBILITY_SCAN_BATCH_SIZE = 5_000;

export function localEpochDaySql(value: string): string {
  return LOCAL_EPOCH_DAY_SQL.replaceAll('%VALUE%', value);
}

function createWordRollupTriggers(db: DatabaseSync): void {
  const dayForNew = localEpochDaySql('NEW.first_seen');
  const dayForOld = localEpochDaySql('OLD.first_seen');

  db.exec(`
    DROP TRIGGER IF EXISTS imm_words_lexical_rollup_insert;
    DROP TRIGGER IF EXISTS imm_words_lexical_rollup_delete;
    DROP TRIGGER IF EXISTS imm_words_lexical_rollup_first_seen_update;

    CREATE TRIGGER imm_words_lexical_rollup_insert
    AFTER INSERT ON imm_words
    WHEN NEW.first_seen IS NOT NULL AND NEW.vocabulary_visible = 1
    BEGIN
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      VALUES (${dayForNew}, 1, CASE WHEN NEW.pos2 = '固有名詞' THEN 0 ELSE 1 END, 0)
      ON CONFLICT(epoch_day) DO UPDATE SET
        word_count = word_count + 1,
        word_count_without_names = word_count_without_names + excluded.word_count_without_names;
    END;

    CREATE TRIGGER imm_words_lexical_rollup_delete
    AFTER DELETE ON imm_words
    WHEN OLD.first_seen IS NOT NULL AND OLD.vocabulary_visible = 1
    BEGIN
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      VALUES (${dayForOld}, -1, CASE WHEN OLD.pos2 = '固有名詞' THEN 0 ELSE -1 END, 0)
      ON CONFLICT(epoch_day) DO UPDATE SET
        word_count = word_count - 1,
        word_count_without_names = word_count_without_names + excluded.word_count_without_names;
      DELETE FROM imm_lexical_daily_rollups
      WHERE epoch_day = ${dayForOld} AND word_count = 0 AND kanji_count = 0;
    END;

    CREATE TRIGGER imm_words_lexical_rollup_first_seen_update
    AFTER UPDATE OF first_seen, pos2, vocabulary_visible ON imm_words
    WHEN OLD.first_seen IS NOT NEW.first_seen
      OR OLD.pos2 IS NOT NEW.pos2
      OR OLD.vocabulary_visible IS NOT NEW.vocabulary_visible
    BEGIN
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      SELECT ${dayForOld}, -1, CASE WHEN OLD.pos2 = '固有名詞' THEN 0 ELSE -1 END, 0
      WHERE OLD.first_seen IS NOT NULL AND OLD.vocabulary_visible = 1
      ON CONFLICT(epoch_day) DO UPDATE SET
        word_count = word_count - 1,
        word_count_without_names = word_count_without_names + excluded.word_count_without_names;
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      SELECT ${dayForNew}, 1, CASE WHEN NEW.pos2 = '固有名詞' THEN 0 ELSE 1 END, 0
      WHERE NEW.first_seen IS NOT NULL AND NEW.vocabulary_visible = 1
      ON CONFLICT(epoch_day) DO UPDATE SET
        word_count = word_count + 1,
        word_count_without_names = word_count_without_names + excluded.word_count_without_names;
      DELETE FROM imm_lexical_daily_rollups
      WHERE word_count = 0 AND kanji_count = 0;
    END;
  `);
}

function createKanjiRollupTriggers(db: DatabaseSync): void {
  const dayForNew = localEpochDaySql('NEW.first_seen');
  const dayForOld = localEpochDaySql('OLD.first_seen');
  db.exec(`
    DROP TRIGGER IF EXISTS imm_kanji_lexical_rollup_insert;
    DROP TRIGGER IF EXISTS imm_kanji_lexical_rollup_delete;
    DROP TRIGGER IF EXISTS imm_kanji_lexical_rollup_first_seen_update;

    CREATE TRIGGER imm_kanji_lexical_rollup_insert
    AFTER INSERT ON imm_kanji WHEN NEW.first_seen IS NOT NULL
    BEGIN
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      VALUES (${dayForNew}, 0, 0, 1)
      ON CONFLICT(epoch_day) DO UPDATE SET kanji_count = kanji_count + 1;
    END;
    CREATE TRIGGER imm_kanji_lexical_rollup_delete
    AFTER DELETE ON imm_kanji WHEN OLD.first_seen IS NOT NULL
    BEGIN
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      VALUES (${dayForOld}, 0, 0, -1)
      ON CONFLICT(epoch_day) DO UPDATE SET kanji_count = kanji_count - 1;
      DELETE FROM imm_lexical_daily_rollups
      WHERE epoch_day = ${dayForOld} AND word_count = 0 AND kanji_count = 0;
    END;
    CREATE TRIGGER imm_kanji_lexical_rollup_first_seen_update
    AFTER UPDATE OF first_seen ON imm_kanji WHEN OLD.first_seen IS NOT NEW.first_seen
    BEGIN
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      SELECT ${dayForOld}, 0, 0, -1 WHERE OLD.first_seen IS NOT NULL
      ON CONFLICT(epoch_day) DO UPDATE SET kanji_count = kanji_count - 1;
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      SELECT ${dayForNew}, 0, 0, 1 WHERE NEW.first_seen IS NOT NULL
      ON CONFLICT(epoch_day) DO UPDATE SET kanji_count = kanji_count + 1;
      DELETE FROM imm_lexical_daily_rollups WHERE word_count = 0 AND kanji_count = 0;
    END;
  `);
}

export function ensureLexicalDailyRollupTables(db: DatabaseSync): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS imm_lexical_daily_rollups(
      epoch_day INTEGER PRIMARY KEY,
      word_count INTEGER NOT NULL DEFAULT 0,
      word_count_without_names INTEGER NOT NULL DEFAULT 0,
      kanji_count INTEGER NOT NULL DEFAULT 0
    );
    INSERT INTO imm_rollup_state(state_key, state_value)
    VALUES ('${LEXICAL_DAILY_ROLLUP_VERSION_KEY}', '0')
    ON CONFLICT(state_key) DO NOTHING;
  `);
  createWordRollupTriggers(db);
  createKanjiRollupTriggers(db);
}

export function areLexicalDailyRollupsReady(db: DatabaseSync): boolean {
  const row = db
    .prepare(`SELECT state_value AS value FROM imm_rollup_state WHERE state_key = ?`)
    .get(LEXICAL_DAILY_ROLLUP_VERSION_KEY) as { value: string } | null;
  return row?.value === LEXICAL_DAILY_ROLLUP_VERSION;
}

export function markLexicalDailyRollupsReady(db: DatabaseSync): void {
  db.prepare(
    `INSERT INTO imm_rollup_state(state_key, state_value)
     VALUES (?, ?)
     ON CONFLICT(state_key) DO UPDATE SET state_value = excluded.state_value`,
  ).run(LEXICAL_DAILY_ROLLUP_VERSION_KEY, LEXICAL_DAILY_ROLLUP_VERSION);
}

/** Rebuild from the first-seen source of truth; run off the UI/main DB thread. */
export function rebuildLexicalDailyRollups(db: DatabaseSync): void {
  let transactionStarted = false;
  try {
    db.exec('BEGIN IMMEDIATE');
    transactionStarted = true;
    const scanVocabulary = db.prepare(
      `SELECT id, word, headword, reading, part_of_speech AS partOfSpeech,
         pos1, pos2, pos3, frequency_rank AS frequencyRank
       FROM imm_words
       WHERE id > ?
       ORDER BY id
       LIMIT ?`,
    );
    const updateVisibility = db.prepare(
      `UPDATE imm_words SET vocabulary_visible = ? WHERE id = ? AND vocabulary_visible IS NOT ?`,
    );
    let lastId = Number.MIN_SAFE_INTEGER;
    for (;;) {
      const vocabularyRows = scanVocabulary.all(
        lastId,
        VOCABULARY_VISIBILITY_SCAN_BATCH_SIZE,
      ) as Array<VocabularyVisibilityRow & { id: number }>;
      if (vocabularyRows.length === 0) break;
      for (const row of vocabularyRows) {
        const visible = isVocabularyStatsRowVisible(row) ? 1 : 0;
        updateVisibility.run(visible, row.id, visible);
      }
      lastId = vocabularyRows[vocabularyRows.length - 1]!.id;
      if (vocabularyRows.length < VOCABULARY_VISIBILITY_SCAN_BATCH_SIZE) break;
    }
    db.exec('DELETE FROM imm_lexical_daily_rollups');
    db.exec(`
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      SELECT ${localEpochDaySql('first_seen')}, COUNT(*),
        SUM(CASE WHEN pos2 = '固有名詞' THEN 0 ELSE 1 END), 0
      FROM imm_words
      WHERE first_seen IS NOT NULL AND vocabulary_visible = 1
      GROUP BY ${localEpochDaySql('first_seen')};
      INSERT INTO imm_lexical_daily_rollups(epoch_day, word_count, word_count_without_names, kanji_count)
      SELECT ${localEpochDaySql('first_seen')}, 0, 0, COUNT(*)
      FROM imm_kanji
      WHERE first_seen IS NOT NULL
      GROUP BY ${localEpochDaySql('first_seen')}
      ON CONFLICT(epoch_day) DO UPDATE SET kanji_count = kanji_count + excluded.kanji_count;
    `);
    markLexicalDailyRollupsReady(db);
    db.exec('COMMIT');
  } catch (error) {
    if (transactionStarted) {
      try {
        db.exec('ROLLBACK');
      } catch {
        // Preserve the rebuild failure; it is the actionable cause.
      }
    }
    throw error;
  }
}

export function getLexicalDailyRollups(db: DatabaseSync): LexicalDailyRollup[] {
  return db
    .prepare(
      `
        SELECT epoch_day AS epochDay, word_count AS wordCount,
          word_count_without_names AS wordCountWithoutNames, kanji_count AS kanjiCount
        FROM imm_lexical_daily_rollups
        ORDER BY epoch_day ASC
      `,
    )
    .all() as LexicalDailyRollup[];
}
