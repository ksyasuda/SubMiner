import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from '../sqlite.js';
import type { DatabaseSync } from '../sqlite.js';
import { ensureSchema } from '../storage.js';
import { deleteSession, deleteSessions, deleteVideo } from '../query-maintenance.js';

const DAY_MS = 86_400_000;
const BASE_MS = 1_700_000_000_000;

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-lexical-removal-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}

/**
 * Seed two episodes of one anime, each with one ended session.
 *
 * `lines` places a word occurrence on a specific day so tests can control which
 * session holds a word's first/last occurrence.
 */
function seed(
  db: DatabaseSync,
  lines: Array<{ session: 1 | 2; wordId: number; dayOffset: number; count?: number }>,
  options: { legacyRows?: boolean } = {},
): void {
  db.exec(`
    INSERT INTO imm_anime(anime_id, normalized_title_key, canonical_title, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'show', 'Show', ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_videos(video_id, video_key, anime_id, canonical_title, source_type, watched, duration_ms, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'v1', 1, 'Ep 1', 1, 1, 1440000, ${BASE_MS}, ${BASE_MS}),
             (2, 'v2', 1, 'Ep 2', 1, 1, 1440000, ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_sessions(session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 's1', 1, '${BASE_MS}', '${BASE_MS + 1000}', 2, ${BASE_MS}, ${BASE_MS}),
             (2, 's2', 2, '${BASE_MS + DAY_MS}', '${BASE_MS + DAY_MS + 1000}', 2, ${BASE_MS}, ${BASE_MS});
  `);

  const insertLine = db.prepare(
    `INSERT INTO imm_subtitle_lines(line_id, session_id, video_id, anime_id, line_index, text, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, ?, 1, ?, ?, ?, ?)`,
  );
  const insertWord = db.prepare(
    `INSERT OR IGNORE INTO imm_words(id, headword, word, reading, part_of_speech, pos1, first_seen, last_seen, frequency)
     VALUES (?, ?, ?, '', 'noun', '名詞', 0, 0, 0)`,
  );
  // `legacyRows` reproduces databases written before the seen_ms column existed,
  // where the timestamp has to be read back off the subtitle line.
  const insertOccurrence = options.legacyRows
    ? db.prepare(
        `INSERT INTO imm_word_line_occurrences(line_id, word_id, occurrence_count) VALUES (?, ?, ?)`,
      )
    : db.prepare(
        `INSERT INTO imm_word_line_occurrences(line_id, word_id, occurrence_count, seen_ms) VALUES (?, ?, ?, ?)`,
      );

  let lineId = 0;
  for (const line of lines) {
    lineId += 1;
    const seenMs = BASE_MS + line.dayOffset * DAY_MS;
    insertLine.run(lineId, line.session, line.session, lineId, `line ${lineId}`, seenMs, seenMs);
    insertWord.run(line.wordId, `語${line.wordId}`, `語${line.wordId}`);
    if (options.legacyRows) {
      insertOccurrence.run(lineId, line.wordId, line.count ?? 1);
    } else {
      insertOccurrence.run(lineId, line.wordId, line.count ?? 1, seenMs);
    }
  }

  // Match what the tracker maintains: aggregates derived from the occurrences.
  db.exec(`
    UPDATE imm_words SET
      frequency = (
        SELECT COALESCE(SUM(o.occurrence_count), 0)
        FROM imm_word_line_occurrences o WHERE o.word_id = imm_words.id
      ),
      first_seen = (
        SELECT MIN(sl.CREATED_DATE) / 1000
        FROM imm_word_line_occurrences o
        JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE o.word_id = imm_words.id
      ),
      last_seen = (
        SELECT MAX(sl.LAST_UPDATE_DATE) / 1000
        FROM imm_word_line_occurrences o
        JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE o.word_id = imm_words.id
      )
  `);
}

function createDb(
  lines: Parameters<typeof seed>[1],
  options: Parameters<typeof seed>[2] = {},
): { db: DatabaseSync; dbPath: string } {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  seed(db, lines, options);
  return { db, dbPath };
}

function readWord(
  db: DatabaseSync,
  wordId: number,
): { frequency: number; firstSeen: number; lastSeen: number } | null {
  return (
    (db
      .prepare(
        'SELECT frequency, first_seen AS firstSeen, last_seen AS lastSeen FROM imm_words WHERE id = ?',
      )
      .get(wordId) as { frequency: number; firstSeen: number; lastSeen: number } | null) ?? null
  );
}

test('deleting a session subtracts only the occurrences it removed', () => {
  const { db, dbPath } = createDb([
    { session: 1, wordId: 10, dayOffset: 0, count: 3 },
    { session: 2, wordId: 10, dayOffset: 1, count: 4 },
  ]);

  try {
    assert.equal(readWord(db, 10)?.frequency, 7);

    deleteSession(db, 1);

    assert.equal(readWord(db, 10)?.frequency, 4, 'only session 1 occurrences are subtracted');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('deleting the session that held a word first drops the word entirely', () => {
  const { db, dbPath } = createDb([
    { session: 1, wordId: 11, dayOffset: 0 },
    { session: 2, wordId: 12, dayOffset: 1 },
  ]);

  try {
    deleteSession(db, 1);

    assert.equal(readWord(db, 11), null, 'word seen only in the deleted session is removed');
    assert.ok(readWord(db, 12), 'word seen elsewhere survives');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('deleting the earliest session moves first_seen forward to the surviving line', () => {
  const { db, dbPath } = createDb([
    { session: 1, wordId: 13, dayOffset: 0 },
    { session: 2, wordId: 13, dayOffset: 5 },
  ]);

  try {
    assert.equal(readWord(db, 13)?.firstSeen, Math.floor(BASE_MS / 1000));

    deleteSession(db, 1);

    const word = readWord(db, 13);
    assert.equal(word?.frequency, 1);
    assert.equal(
      word?.firstSeen,
      Math.floor((BASE_MS + 5 * DAY_MS) / 1000),
      'first_seen advances to the remaining occurrence',
    );
    assert.equal(word?.lastSeen, Math.floor((BASE_MS + 5 * DAY_MS) / 1000));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('deleting the latest session moves last_seen back to the surviving line', () => {
  const { db, dbPath } = createDb([
    { session: 1, wordId: 14, dayOffset: 0 },
    { session: 2, wordId: 14, dayOffset: 5 },
  ]);

  try {
    deleteSession(db, 2);

    const word = readWord(db, 14);
    assert.equal(word?.frequency, 1);
    assert.equal(word?.lastSeen, Math.floor(BASE_MS / 1000), 'last_seen falls back to session 1');
    assert.equal(word?.firstSeen, Math.floor(BASE_MS / 1000));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('deleting an interior occurrence leaves the surrounding extremes untouched', () => {
  // Session 2 carries the middle occurrence; sessions bracket it in time.
  const { db, dbPath } = createDb([
    { session: 1, wordId: 15, dayOffset: 0 },
    { session: 2, wordId: 15, dayOffset: 3 },
    { session: 1, wordId: 15, dayOffset: 9 },
  ]);

  try {
    deleteSessions(db, [2]);

    const word = readWord(db, 15);
    assert.equal(word?.frequency, 2);
    assert.equal(word?.firstSeen, Math.floor(BASE_MS / 1000));
    assert.equal(word?.lastSeen, Math.floor((BASE_MS + 9 * DAY_MS) / 1000));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('a stored frequency that has drifted low is repaired instead of dropping a live word', () => {
  const { db, dbPath } = createDb([
    { session: 1, wordId: 16, dayOffset: 0, count: 5 },
    { session: 2, wordId: 16, dayOffset: 4, count: 5 },
  ]);

  try {
    // Simulate drift: the stored total is lower than the occurrences justify, so
    // naive subtraction would take the word to zero while rows still reference it.
    db.prepare('UPDATE imm_words SET frequency = 5 WHERE id = ?').run(16);

    deleteSession(db, 1);

    const word = readWord(db, 16);
    assert.ok(word, 'word with surviving occurrences is not deleted');
    assert.equal(word?.frequency, 5, 'frequency is recomputed from the surviving occurrences');
    assert.equal(word?.firstSeen, Math.floor((BASE_MS + 4 * DAY_MS) / 1000));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('deleting a video subtracts every occurrence carried by its lines', () => {
  const { db, dbPath } = createDb([
    { session: 1, wordId: 17, dayOffset: 0, count: 2 },
    { session: 1, wordId: 17, dayOffset: 1, count: 3 },
    { session: 2, wordId: 17, dayOffset: 2, count: 4 },
    { session: 2, wordId: 18, dayOffset: 2, count: 1 },
  ]);

  try {
    deleteVideo(db, 1);

    assert.equal(readWord(db, 17)?.frequency, 4, 'both lines from video 1 are subtracted');
    assert.equal(readWord(db, 18)?.frequency, 1, 'untouched video keeps its counts');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('occurrence rows written before the seen_ms column still resolve their dates', () => {
  const { db, dbPath } = createDb(
    [
      { session: 1, wordId: 20, dayOffset: 0 },
      { session: 2, wordId: 20, dayOffset: 6 },
    ],
    { legacyRows: true },
  );

  try {
    assert.equal(
      (
        db
          .prepare('SELECT COUNT(*) AS total FROM imm_word_line_occurrences WHERE seen_ms IS NULL')
          .get() as { total: number }
      ).total,
      2,
      'precondition: rows carry no denormalised timestamp',
    );

    deleteSession(db, 1);

    const word = readWord(db, 20);
    assert.equal(word?.frequency, 1);
    assert.equal(
      word?.firstSeen,
      Math.floor((BASE_MS + 6 * DAY_MS) / 1000),
      'falls back to the subtitle line timestamp',
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('upgrading an older database backfills seen_ms from the subtitle lines', () => {
  const { db, dbPath } = createDb(
    [
      { session: 1, wordId: 21, dayOffset: 0 },
      { session: 2, wordId: 21, dayOffset: 2 },
    ],
    { legacyRows: true },
  );

  try {
    // Re-run ensureSchema the way a pre-19 database would be opened.
    db.exec('DELETE FROM imm_schema_version');
    db.exec(`INSERT INTO imm_schema_version(schema_version, applied_at_ms) VALUES (18, '0')`);
    ensureSchema(db);

    const rows = db
      .prepare('SELECT line_id AS lineId, seen_ms AS seenMs FROM imm_word_line_occurrences')
      .all() as Array<{ lineId: number; seenMs: number | null }>;
    assert.equal(rows.length, 2);
    for (const row of rows) {
      assert.ok(row.seenMs, `line ${row.lineId} should have a backfilled timestamp`);
    }
    assert.deepEqual(
      rows.map((row) => row.seenMs).sort((a, b) => Number(a) - Number(b)),
      [BASE_MS, BASE_MS + 2 * DAY_MS],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
