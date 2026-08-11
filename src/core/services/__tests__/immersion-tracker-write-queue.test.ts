import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { DatabaseSync } from '../immersion-tracker/sqlite';

type ImmersionTrackerService = import('../immersion-tracker-service').ImmersionTrackerService;
type ImmersionTrackerServiceCtor =
  typeof import('../immersion-tracker-service').ImmersionTrackerService;

let trackerCtor: ImmersionTrackerServiceCtor | null = null;

async function loadTrackerCtor(): Promise<ImmersionTrackerServiceCtor> {
  if (trackerCtor) return trackerCtor;
  const mod = await import('../immersion-tracker-service');
  trackerCtor = mod.ImmersionTrackerService;
  return trackerCtor;
}

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-write-queue-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) return;
  fs.rmSync(dir, { recursive: true, force: true });
}

interface TrackerInternals {
  db: DatabaseSync;
  queue: unknown[];
  recordWrite: (write: Record<string, unknown>) => void;
  mergeAnime: (targetAnimeId: number, sourceAnimeIds: number[]) => Promise<unknown>;
  moveVideoToAnime: (videoId: number, targetAnimeId: number) => Promise<unknown>;
}

function seedTwoEntries(db: DatabaseSync): void {
  db.exec(`
    INSERT INTO imm_anime (anime_id, normalized_title_key, canonical_title, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'show', 'Show', 1000, 1000), (2, 'show season 1', 'Show Season 1', 1000, 1000);
    INSERT INTO imm_videos (video_id, video_key, canonical_title, anime_id, source_type, watched, duration_ms, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'local:/tmp/a.mkv', 'A', 1, 1, 0, 1440000, 1000, 1000),
             (2, 'local:/tmp/b.mkv', 'B', 2, 1, 0, 1440000, 1000, 1000);
    INSERT INTO imm_sessions (session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, active_watched_ms, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'drain-session', 2, '1000', '2000', 2, 1000, 1000, 2000);
  `);
}

function queueSubtitleLines(tracker: TrackerInternals, count: number): void {
  for (let index = 0; index < count; index += 1) {
    tracker.recordWrite({
      kind: 'subtitleLine',
      sessionId: 1,
      videoId: 2,
      lineIndex: index,
      segmentStartMs: index * 1000,
      segmentEndMs: index * 1000 + 900,
      text: `line ${index}`,
      wordOccurrences: [],
      kanjiOccurrences: [],
      firstSeen: 1000,
      lastSeen: 2000,
    });
  }
}

function countLinesForAnime(db: DatabaseSync, animeId: number): number {
  const row = db
    .prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines WHERE anime_id = ?')
    .get(animeId) as { total: number };
  return Number(row.total);
}

/**
 * Both entry points rebuild the lifetime summaries, which recompute from the
 * database. A single flushNow() only writes one batch off the front of the
 * queue, so anything past `batchSize` would still be unwritten when the rebuild
 * reads.
 */
test('mergeAnime drains a queue larger than one batch before rebuilding summaries', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath, policy: { batchSize: 2 } });
    const internals = tracker as unknown as TrackerInternals;

    seedTwoEntries(internals.db);
    queueSubtitleLines(internals, 8);
    assert.ok(internals.queue.length > 2, 'expected more queued writes than one batch');

    await internals.mergeAnime(1, [2]);

    assert.equal(internals.queue.length, 0);
    assert.equal(countLinesForAnime(internals.db, 1), 8);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('moveVideoToAnime drains a queue larger than one batch before rebuilding summaries', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath, policy: { batchSize: 2 } });
    const internals = tracker as unknown as TrackerInternals;

    seedTwoEntries(internals.db);
    queueSubtitleLines(internals, 8);

    await internals.moveVideoToAnime(2, 1);

    assert.equal(internals.queue.length, 0);
    assert.equal(countLinesForAnime(internals.db, 1), 8);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});
