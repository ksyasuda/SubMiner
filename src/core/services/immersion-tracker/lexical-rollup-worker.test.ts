import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { areLexicalDailyRollupsReady, getLexicalDailyRollups } from './lexical-rollups';
import { executeLexicalRollupBackfillTask } from './lexical-rollup-worker';
import { Database } from './sqlite';
import { ensureSchema } from './storage';

test('lexical rollup backfill materializes pre-existing vocabulary off the caller DB connection', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-lexical-rollup-worker-'));
  const dbPath = path.join(directory, 'immersion.sqlite');
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES (?, ?, ?, ?, ?, 1)`,
    ).run('犬', '犬', 'いぬ', 1_700_000_000, 1_700_000_000);
    db.exec('DELETE FROM imm_lexical_daily_rollups');
    db.prepare(`UPDATE imm_rollup_state SET state_value = '0' WHERE state_key = ?`).run(
      'lexical_daily_rollups_ready',
    );

    executeLexicalRollupBackfillTask(dbPath);

    assert.equal(areLexicalDailyRollupsReady(db), true);
    assert.equal(getLexicalDailyRollups(db)[0]?.wordCount, 1);
  } finally {
    db.close();
    fs.rmSync(directory, { recursive: true, force: true });
  }
});
