import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  resolveVocabularySummaryWorkerPath,
  VocabularySummaryWorkerRuntime,
} from './vocabulary-summary-worker-runtime';
import { Database } from './sqlite';
import { applyPragmas, ensureSchema } from './storage';

test('vocabulary summary worker reads the database from a separate connection', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-vocabulary-summary-worker-'));
  const dbPath = path.join(tempDir, 'immersion.sqlite');
  const runtime = new VocabularySummaryWorkerRuntime();
  const db = new Database(dbPath);

  try {
    applyPragmas(db);
    ensureSchema(db);
    db.prepare(
      `
      INSERT INTO imm_words (
        headword, word, reading, part_of_speech, pos1, pos2, pos3,
        first_seen, last_seen, frequency
      ) VALUES ('猫', '猫', 'ねこ', 'noun', '名詞', '一般', '', 1, 1, 1)
    `,
    ).run();
    db.close();

    const summary = await runtime.run(dbPath, new Set(['猫']));

    assert.equal(summary.uniqueWords, 1);
    assert.equal(summary.knownWordCount, 1);
  } finally {
    runtime.destroy();
    try {
      db.close();
    } catch {
      // The worker needs the setup connection closed before it starts.
    }
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('vocabulary summary worker module resolves in the current layout', () => {
  const workerPath = resolveVocabularySummaryWorkerPath();
  assert.ok(workerPath, 'expected the vocabulary summary worker module to resolve');
  assert.ok(workerPath.endsWith(__filename.endsWith('.ts') ? '.ts' : '.js'));
});

test('vocabulary summary worker never falls back to the caller thread', async () => {
  const runtime = new VocabularySummaryWorkerRuntime({
    resolveWorkerPath: () => null,
    warn: () => {},
  });

  try {
    await assert.rejects(
      runtime.run('/tmp/subminer-summary-worker-not-used.sqlite', null),
      /worker unavailable/i,
    );
  } finally {
    runtime.destroy();
  }
});
