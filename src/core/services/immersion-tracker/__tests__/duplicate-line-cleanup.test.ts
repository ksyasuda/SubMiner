import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database } from '../sqlite.js';
import type { DatabaseSync } from '../sqlite.js';
import { ensureSchema } from '../storage.js';
import { cleanupDuplicateSubtitleLines } from '../duplicate-line-cleanup.js';

const DAY_MS = 86_400_000;
const BASE_MS = 1_700_000_000_000;
const WORD_ID = 1;

interface SeedLine {
  session: number;
  text: string;
  startMs: number;
  endMs: number;
  /** Recording wall-clock, i.e. what the lookback window filters on. */
  createdMs?: number;
}

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-duplicate-line-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}

/** One episode, two sessions of it, and one word occurrence per seeded line. */
function seed(db: DatabaseSync, lines: SeedLine[]): void {
  db.exec(`
    INSERT INTO imm_anime(anime_id, normalized_title_key, canonical_title, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'show', 'Show', ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_videos(video_id, video_key, anime_id, canonical_title, source_type, watched, duration_ms, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'v1', 1, 'Ep 1', 1, 1, 1440000, ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_sessions(session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 's1', 1, '${BASE_MS}', '${BASE_MS + 1000}', 2, ${BASE_MS}, ${BASE_MS}),
             (2, 's2', 1, '${BASE_MS + DAY_MS}', '${BASE_MS + DAY_MS + 1000}', 2, ${BASE_MS}, ${BASE_MS});
    INSERT INTO imm_words(id, headword, word, reading, part_of_speech, pos1, first_seen, last_seen, frequency)
      VALUES (${WORD_ID}, '飛び上がる', '飛び上がる', '', 'verb', '動詞', ${Math.floor(BASE_MS / 1000)}, ${Math.floor(BASE_MS / 1000)}, 0);
  `);

  const insertLine = db.prepare(
    `INSERT INTO imm_subtitle_lines(
       line_id, session_id, video_id, anime_id, line_index,
       segment_start_ms, segment_end_ms, text, CREATED_DATE, LAST_UPDATE_DATE)
     VALUES (?, ?, 1, 1, ?, ?, ?, ?, ?, ?)`,
  );
  const insertOccurrence = db.prepare(
    `INSERT INTO imm_word_line_occurrences(line_id, word_id, occurrence_count, seen_ms)
     VALUES (?, ?, 1, ?)`,
  );

  lines.forEach((line, index) => {
    const lineId = index + 1;
    const createdMs = line.createdMs ?? BASE_MS;
    insertLine.run(
      lineId,
      line.session,
      lineId,
      line.startMs,
      line.endMs,
      line.text,
      createdMs,
      createdMs,
    );
    insertOccurrence.run(lineId, WORD_ID, createdMs);
  });

  db.exec(`
    UPDATE imm_words SET frequency = (
      SELECT COALESCE(SUM(o.occurrence_count), 0)
      FROM imm_word_line_occurrences o WHERE o.word_id = imm_words.id
    )
  `);
}

function createDb(lines: SeedLine[]): { db: DatabaseSync; dbPath: string } {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);
  ensureSchema(db);
  seed(db, lines);
  return { db, dbPath };
}

/** A typeset line mpv reported once per animation frame. */
function karaokeFrames(
  session: number,
  text: string,
  startMs: number,
  frames: number,
  frameMs: number,
): SeedLine[] {
  return Array.from({ length: frames }, (_, index) => ({
    session,
    text,
    startMs: startMs + index * frameMs,
    endMs: startMs + (index + 1) * frameMs,
  }));
}

function countLines(db: DatabaseSync): number {
  return (db.prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines').get() as { total: number })
    .total;
}

function wordFrequency(db: DatabaseSync): number {
  const row = db.prepare('SELECT frequency FROM imm_words WHERE id = ?').get(WORD_ID) as {
    frequency: number;
  } | null;
  return row?.frequency ?? 0;
}

test('a karaoke burst collapses to one line and gives back its word counts', () => {
  const { db, dbPath } = createDb([
    ...karaokeFrames(1, '飛び上がる', 10_000, 40, 40),
    { session: 1, text: 'おはよう', startMs: 20_000, endMs: 22_000 },
  ]);

  try {
    const summary = cleanupDuplicateSubtitleLines(db);

    assert.equal(summary.burstGroups, 1);
    assert.equal(summary.removedLines, 39);
    assert.equal(summary.removedWordOccurrences, 39);
    assert.equal(countLines(db), 2);
    assert.equal(wordFrequency(db), 2);

    // The surviving line covers the whole run, the way the parsed cue would.
    const kept = db
      .prepare(
        'SELECT segment_start_ms AS startMs, segment_end_ms AS endMs FROM imm_subtitle_lines WHERE line_id = 1',
      )
      .get() as { startMs: number; endMs: number };
    assert.equal(kept.startMs, 10_000);
    assert.equal(kept.endMs, 10_000 + 40 * 40);

    assert.equal(summary.samples.length, 1);
    assert.equal(summary.samples[0]!.text, '飛び上がる');
    assert.equal(summary.samples[0]!.frames, 40);
    assert.equal(summary.samples[0]!.videoTitle, 'Ep 1');
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ordinary repeated dialogue survives', () => {
  // Six contiguous `飛び上がる`, each held for a normal beat rather than a frame.
  const lines = Array.from({ length: 6 }, (_, index) => ({
    session: 1,
    text: '飛び上がる',
    startMs: 5_000 + index * 800,
    endMs: 5_000 + (index + 1) * 800,
  }));
  const { db, dbPath } = createDb(lines);

  try {
    const summary = cleanupDuplicateSubtitleLines(db);

    assert.equal(summary.burstGroups, 0);
    assert.equal(summary.removedLines, 0);
    assert.equal(countLines(db), 6);
    assert.equal(wordFrequency(db), 6);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('a short run below the threshold survives', () => {
  const { db, dbPath } = createDb(karaokeFrames(1, '飛び上がる', 1_000, 4, 40));

  try {
    const summary = cleanupDuplicateSubtitleLines(db);

    assert.equal(summary.burstGroups, 0);
    assert.equal(countLines(db), 4);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('the same line in a rewatch session is never merged into the first watch', () => {
  const { db, dbPath } = createDb([
    ...karaokeFrames(1, '飛び上がる', 10_000, 6, 40),
    ...karaokeFrames(2, '飛び上がる', 10_000, 6, 40),
  ]);

  try {
    const summary = cleanupDuplicateSubtitleLines(db);

    assert.equal(summary.burstGroups, 2);
    assert.equal(summary.removedLines, 10);
    // One surviving line per session, not one across both.
    assert.equal(countLines(db), 2);
    assert.equal(wordFrequency(db), 2);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('a gap between runs splits them', () => {
  const { db, dbPath } = createDb([
    ...karaokeFrames(1, '飛び上がる', 10_000, 6, 40),
    ...karaokeFrames(1, '飛び上がる', 60_000, 6, 40),
  ]);

  try {
    const summary = cleanupDuplicateSubtitleLines(db);

    assert.equal(summary.burstGroups, 2);
    assert.equal(countLines(db), 2);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('a dry run reports what an apply would do and writes nothing', () => {
  const { db, dbPath } = createDb(karaokeFrames(1, '飛び上がる', 10_000, 40, 40));

  try {
    const preview = cleanupDuplicateSubtitleLines(db, { dryRun: true });

    assert.equal(preview.dryRun, true);
    assert.equal(preview.removedLines, 39);
    assert.equal(countLines(db), 40);
    assert.equal(wordFrequency(db), 40);

    const applied = cleanupDuplicateSubtitleLines(db);
    assert.equal(applied.removedLines, preview.removedLines);
    assert.equal(applied.removedWordOccurrences, preview.removedWordOccurrences);
    assert.equal(countLines(db), 1);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('the lookback window leaves older bursts alone', () => {
  const recentMs = BASE_MS;
  const oldMs = BASE_MS - 40 * DAY_MS;
  const { db, dbPath } = createDb([
    ...karaokeFrames(1, '飛び上がる', 10_000, 6, 40).map((line) => ({
      ...line,
      createdMs: oldMs,
    })),
    ...karaokeFrames(2, '飛び上がる', 10_000, 6, 40).map((line) => ({
      ...line,
      createdMs: recentMs,
    })),
  ]);

  globalThis.__subminerTestNowMs = BASE_MS;
  try {
    const summary = cleanupDuplicateSubtitleLines(db, { lookbackDays: 30 });

    assert.equal(summary.lookbackDays, 30);
    assert.equal(summary.scannedLines, 6);
    assert.equal(summary.burstGroups, 1);
    assert.equal(summary.removedLines, 5);
    // Six untouched old frames plus the one surviving recent line.
    assert.equal(countLines(db), 7);
    assert.equal(wordFrequency(db), 7);
  } finally {
    globalThis.__subminerTestNowMs = undefined;
    db.close();
    cleanupDbPath(dbPath);
  }
});
