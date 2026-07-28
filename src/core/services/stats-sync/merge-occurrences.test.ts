import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from '../immersion-tracker/sqlite';
import { ensureSchema } from '../immersion-tracker/storage';
import { mergeSnapshotIntoDb } from './merge';

const BASE_MS = 1_700_000_000_000;
const DAY_MS = 86_400_000;

/**
 * Build a database holding one anime, one episode, one ended session and a
 * single word occurrence.
 *
 * `legacyOccurrences` reproduces a peer whose schema predates the denormalised
 * `seen_ms` column, so the merge has to recover the timestamp from the line.
 */
function buildDb(
  dir: string,
  name: string,
  options: { word: string; seenMs: number; legacyOccurrences: boolean },
): string {
  const dbPath = path.join(dir, name);
  const db = new Database(dbPath);
  ensureSchema(db);
  db.exec(`
    INSERT INTO imm_anime(anime_id, normalized_title_key, canonical_title, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'key-${name}', 'Show ${name}', ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_videos(video_id, video_key, anime_id, canonical_title, source_type, watched, duration_ms, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'video-${name}', 1, 'Episode 1', 1, 1, 1440000, ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_sessions(session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'uuid-${name}', 1, '${options.seenMs}', '${options.seenMs + 1000}', 2, ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_subtitle_lines(line_id, session_id, video_id, anime_id, line_index, text, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 1, 1, 1, 0, 'line', ${options.seenMs}, ${options.seenMs});
    INSERT INTO imm_words(id, headword, word, reading, part_of_speech, pos1, first_seen, last_seen, frequency)
      VALUES (1, '${options.word}', '${options.word}', '', 'noun', '名詞',
              ${Math.floor(options.seenMs / 1000)}, ${Math.floor(options.seenMs / 1000)}, 1);
  `);
  db.exec(
    options.legacyOccurrences
      ? 'INSERT INTO imm_word_line_occurrences(line_id, word_id, occurrence_count) VALUES (1, 1, 1)'
      : `INSERT INTO imm_word_line_occurrences(line_id, word_id, occurrence_count, seen_ms) VALUES (1, 1, 1, ${options.seenMs})`,
  );
  db.close();
  return dbPath;
}

function mergedOccurrences(dbPath: string): Array<{ word: string; seenMs: number | null }> {
  const db = new Database(dbPath);
  try {
    return db
      .prepare(
        `SELECT w.word AS word, o.seen_ms AS seenMs
         FROM imm_word_line_occurrences o
         JOIN imm_words w ON w.id = o.word_id
         ORDER BY w.word`,
      )
      .all() as Array<{ word: string; seenMs: number | null }>;
  } finally {
    db.close();
  }
}

for (const legacyOccurrences of [false, true]) {
  const label = legacyOccurrences ? 'a peer predating the seen_ms column' : 'a current peer';

  test(`sync merge carries occurrence timestamps in from ${label}`, () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-merge-occurrences-test-'));
    try {
      const local = buildDb(dir, 'local.sqlite', {
        word: '猫',
        seenMs: BASE_MS,
        legacyOccurrences: false,
      });
      const remoteSeenMs = BASE_MS + 5 * DAY_MS;
      const remote = buildDb(dir, 'remote.sqlite', {
        word: '犬',
        seenMs: remoteSeenMs,
        legacyOccurrences,
      });

      mergeSnapshotIntoDb(local, remote);

      const rows = mergedOccurrences(local);
      assert.equal(rows.length, 2, 'local and remote occurrences both survive the merge');
      for (const row of rows) {
        assert.ok(row.seenMs, `${row.word} must carry a timestamp so aggregates stay index-only`);
      }
      assert.equal(
        Number(rows.find((row) => row.word === '犬')?.seenMs),
        remoteSeenMs,
        "the remote line's own timestamp is preserved rather than stamped with merge time",
      );
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
}
