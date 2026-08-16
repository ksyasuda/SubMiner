import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  LexicalRollupWorkerRuntime,
  resolveLexicalRollupWorkerPath,
} from './lexical-rollup-worker-runtime';
import { areLexicalDailyRollupsReady } from './lexical-rollups';
import { Database } from './sqlite';
import { applyPragmas, ensureSchema } from './storage';

test('lexical rollup worker backfills without using the tracker connection', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-lexical-rollup-runtime-'));
  const dbPath = path.join(directory, 'immersion.sqlite');
  const runtime = new LexicalRollupWorkerRuntime();
  const db = new Database(dbPath);

  try {
    applyPragmas(db);
    ensureSchema(db);
    db.prepare(
      `INSERT INTO imm_words(headword, word, reading, first_seen, last_seen, frequency)
       VALUES ('鳥', '鳥', 'とり', 1700000000, 1700000000, 1)`,
    ).run();
    db.exec('DELETE FROM imm_lexical_daily_rollups');
    db.prepare(`UPDATE imm_rollup_state SET state_value = '0' WHERE state_key = ?`).run(
      'lexical_daily_rollups_ready',
    );
    db.close();

    await runtime.run(dbPath);

    const checkDb = new Database(dbPath);
    try {
      assert.equal(areLexicalDailyRollupsReady(checkDb), true);
    } finally {
      checkDb.close();
    }
  } finally {
    runtime.destroy();
    try {
      db.close();
    } catch {
      // Closed before the worker starts.
    }
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('lexical rollup worker module resolves in the current layout', () => {
  const workerPath = resolveLexicalRollupWorkerPath();
  assert.ok(workerPath, 'expected the lexical rollup worker module to resolve');
  assert.ok(workerPath.endsWith(__filename.endsWith('.ts') ? '.ts' : '.js'));
});

test('lexical rollup worker leaves a backfill pending when no worker can start', async () => {
  const runtime = new LexicalRollupWorkerRuntime({
    resolveWorkerPath: () => null,
    warn: () => {},
  } as never);

  try {
    await assert.doesNotReject(runtime.run('/tmp/not-used.sqlite'));
  } finally {
    runtime.destroy();
  }
});
