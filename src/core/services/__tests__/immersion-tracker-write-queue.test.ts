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
  rebuildLifetimeSummaries: () => Promise<unknown>;
  flushNow: () => void;
}

test('mergeAnime fails closed when queued writes cannot drain', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath, policy: { batchSize: 2 } });
    const internals = tracker as unknown as TrackerInternals;
    seedTwoEntries(internals.db);
    queueSubtitleLines(internals, 1);
    internals.flushNow = () => {};

    await assert.rejects(internals.mergeAnime(1, [2]), /queue did not drain/i);

    assert.deepEqual(
      internals.db
        .prepare('SELECT anime_id AS animeId FROM imm_anime ORDER BY anime_id')
        .all()
        .map((row) => (row as { animeId: number }).animeId),
      [1, 2],
    );
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('moveVideoToAnime fails closed when queued writes cannot drain', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath, policy: { batchSize: 2 } });
    const internals = tracker as unknown as TrackerInternals;
    seedTwoEntries(internals.db);
    queueSubtitleLines(internals, 1);
    internals.flushNow = () => {};

    await assert.rejects(internals.moveVideoToAnime(2, 1), /queue did not drain/i);
    assert.equal(
      (
        internals.db
          .prepare('SELECT anime_id AS animeId FROM imm_videos WHERE video_id = 2')
          .get() as {
          animeId: number;
        }
      ).animeId,
      2,
    );
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('rebuildLifetimeSummaries fails closed when queued writes cannot drain', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath, policy: { batchSize: 2 } });
    const internals = tracker as unknown as TrackerInternals;
    seedTwoEntries(internals.db);
    queueSubtitleLines(internals, 1);
    internals.flushNow = () => {};

    await assert.rejects(internals.rebuildLifetimeSummaries(), /queue did not drain/i);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

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

/**
 * Queued last so it sits past the first batch. Lifetime `total_lines_seen`
 * reads this counter, not a COUNT over imm_subtitle_lines, so the rebuilt
 * summary only reflects the session once the queue is drained all the way.
 */
function queueTelemetry(tracker: TrackerInternals, linesSeen: number): void {
  tracker.recordWrite({
    kind: 'telemetry',
    sessionId: 1,
    sampleMs: 3000,
    lastMediaMs: 3000,
    totalWatchedMs: 4000,
    activeWatchedMs: 3500,
    linesSeen,
    tokensSeen: linesSeen * 5,
    cardsMined: 2,
    lookupCount: 0,
    lookupHits: 0,
    yomitanLookupCount: 0,
    pauseCount: 0,
    pauseMs: 0,
    seekForwardCount: 0,
    seekBackwardCount: 0,
    mediaBufferEvents: 0,
  });
}

function lifetimeForAnime(
  db: DatabaseSync,
  animeId: number,
): { linesSeen: number; activeMs: number; cards: number } | null {
  const row = db
    .prepare(
      `SELECT total_lines_seen AS linesSeen, total_active_ms AS activeMs, total_cards AS cards
       FROM imm_lifetime_anime WHERE anime_id = ?`,
    )
    .get(animeId) as { linesSeen: number; activeMs: number; cards: number } | undefined;
  return row
    ? { linesSeen: Number(row.linesSeen), activeMs: Number(row.activeMs), cards: Number(row.cards) }
    : null;
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
    queueTelemetry(internals, 8);
    assert.ok(internals.queue.length > 2, 'expected more queued writes than one batch');

    await internals.mergeAnime(1, [2]);

    assert.equal(internals.queue.length, 0);
    assert.equal(countLinesForAnime(internals.db, 1), 8);
    // The surviving entry's summary was rebuilt from the fully drained queue.
    const lifetime = lifetimeForAnime(internals.db, 1);
    assert.equal(lifetime?.linesSeen, 8);
    assert.equal(lifetime?.activeMs, 3500);
    assert.equal(lifetime?.cards, 2);
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
    queueTelemetry(internals, 8);

    await internals.moveVideoToAnime(2, 1);

    assert.equal(internals.queue.length, 0);
    assert.equal(countLinesForAnime(internals.db, 1), 8);
    const lifetime = lifetimeForAnime(internals.db, 1);
    assert.equal(lifetime?.linesSeen, 8);
    assert.equal(lifetime?.activeMs, 3500);
    assert.equal(lifetime?.cards, 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});
