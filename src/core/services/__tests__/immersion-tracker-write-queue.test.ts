import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { DatabaseSync } from '../immersion-tracker/sqlite';

interface TrackerInternals {
  db: DatabaseSync;
  queue: unknown[];
  recordWrite: (write: Record<string, unknown>) => void;
  deleteSession: (sessionId: number) => Promise<void>;
  mergeAnime: (targetAnimeId: number, sourceAnimeIds: number[]) => Promise<unknown>;
  moveVideoToAnime: (videoId: number, targetAnimeId: number) => Promise<unknown>;
  rebuildLifetimeSummaries: () => Promise<unknown>;
  reassignAnimeAnilist: (animeId: number, info: { anilistId: number }) => Promise<void>;
  flushNow: () => void;
  writeLock: { locked: boolean };
}

interface TrackerHarness {
  tracker: TrackerInternals;
  deleteRunnerCalls: () => number;
}

/** Runs `fn` against a fresh tracker (batchSize 2) seeded with two library entries. */
async function withTracker(fn: (harness: TrackerHarness) => Promise<void>): Promise<void> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-write-queue-test-'));
  const { ImmersionTrackerService } = await import('../immersion-tracker-service');
  let deleteRunnerCalls = 0;
  const service = new ImmersionTrackerService(
    { dbPath: path.join(dir, 'immersion.sqlite'), policy: { batchSize: 2 } },
    {
      runDeleteMaintenanceTask: async () => {
        deleteRunnerCalls += 1;
      },
    },
  );
  try {
    const tracker = service as unknown as TrackerInternals;
    seedTwoEntries(tracker.db);
    await fn({ tracker, deleteRunnerCalls: () => deleteRunnerCalls });
  } finally {
    service.destroy();
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function seedTwoEntries(db: DatabaseSync): void {
  db.exec(`
    INSERT INTO imm_anime (anime_id, normalized_title_key, canonical_title, anilist_id, CREATED_DATE, LAST_UPDATE_DATE)
      VALUES (1, 'show', 'Show', NULL, 1000, 1000), (2, 'show season 1', 'Show Season 1', 123, 1000, 1000);
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

const SNAPSHOT_TABLES = [
  'imm_anime',
  'imm_videos',
  'imm_sessions',
  'imm_subtitle_lines',
  'imm_session_telemetry',
  'imm_lifetime_global',
  'imm_lifetime_anime',
  'imm_lifetime_media',
  'imm_lifetime_applied_sessions',
];

function snapshotDb(db: DatabaseSync): Record<string, unknown[]> {
  return Object.fromEntries(
    SNAPSHOT_TABLES.map((table) => [table, db.prepare(`SELECT * FROM ${table} ORDER BY 1`).all()]),
  );
}

const guardedOperations: Array<{
  name: string;
  run: (tracker: TrackerInternals) => Promise<unknown>;
}> = [
  { name: 'deleteSession', run: (tracker) => tracker.deleteSession(1) },
  { name: 'mergeAnime', run: (tracker) => tracker.mergeAnime(1, [2]) },
  { name: 'moveVideoToAnime', run: (tracker) => tracker.moveVideoToAnime(2, 1) },
  { name: 'rebuildLifetimeSummaries', run: (tracker) => tracker.rebuildLifetimeSummaries() },
  {
    // Anime 2 already owns AniList id 123, so this would otherwise resolve a conflict.
    name: 'reassignAnimeAnilist',
    run: (tracker) => tracker.reassignAnimeAnilist(1, { anilistId: 123 }),
  },
];

for (const operation of guardedOperations) {
  test(`${operation.name} fails closed without touching the database when queued writes cannot drain`, async () => {
    await withTracker(async ({ tracker, deleteRunnerCalls }) => {
      queueSubtitleLines(tracker, 1);
      tracker.flushNow = () => {};
      const before = snapshotDb(tracker.db);

      await assert.rejects(operation.run(tracker), /queue did not drain/i);

      assert.deepEqual(snapshotDb(tracker.db), before);
      assert.equal(deleteRunnerCalls(), 0);
      assert.equal(tracker.writeLock.locked, false);
    });
  });
}

/**
 * Both entry points must see a settled database before changing episode
 * ownership. A single flushNow() only writes one batch off the front of the
 * queue, so anything past `batchSize` would still be unwritten when the merge
 * repoints rows.
 */
const repointingOperations: Array<{
  name: string;
  run: (tracker: TrackerInternals) => Promise<unknown>;
}> = [
  { name: 'mergeAnime', run: (tracker) => tracker.mergeAnime(1, [2]) },
  { name: 'moveVideoToAnime', run: (tracker) => tracker.moveVideoToAnime(2, 1) },
];

for (const operation of repointingOperations) {
  test(`${operation.name} drains a queue larger than one batch before repointing rows`, async () => {
    await withTracker(async ({ tracker }) => {
      queueSubtitleLines(tracker, 8);
      queueTelemetry(tracker, 8);
      assert.ok(tracker.queue.length > 2, 'expected more queued writes than one batch');

      await operation.run(tracker);

      assert.equal(tracker.queue.length, 0);
      // Every queued line and the trailing telemetry sample landed on the surviving entry.
      const lines = tracker.db
        .prepare('SELECT COUNT(*) AS total FROM imm_subtitle_lines WHERE anime_id = 1')
        .get() as { total: number };
      assert.equal(Number(lines.total), 8);
      const telemetry = tracker.db
        .prepare(
          `SELECT lines_seen AS linesSeen FROM imm_session_telemetry
           WHERE session_id = 1 ORDER BY sample_ms DESC, telemetry_id DESC LIMIT 1`,
        )
        .get() as { linesSeen: number } | undefined;
      assert.equal(Number(telemetry?.linesSeen), 8);
    });
  });
}
