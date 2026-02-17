import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";
import { ImmersionTrackerService } from "./immersion-tracker-service";

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "subminer-immersion-test-"));
  return path.join(dir, "immersion.sqlite");
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (fs.existsSync(dir)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test("startSession generates UUID-like session identifiers", () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    tracker = new ImmersionTrackerService({ dbPath });
    tracker.handleMediaChange("/tmp/episode.mkv", "Episode");

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const db = new DatabaseSync(dbPath);
    const row = db
      .prepare("SELECT session_uuid FROM imm_sessions LIMIT 1")
      .get() as { session_uuid: string } | null;
    db.close();

    assert.equal(typeof row?.session_uuid, "string");
    assert.equal(row?.session_uuid?.startsWith("session-"), false);
    assert.ok(/^[0-9a-fA-F-]{36}$/.test(row?.session_uuid || ""));
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test("destroy finalizes session with a single telemetry flush path", () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    tracker = new ImmersionTrackerService({ dbPath });
    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
    };

    let flushCalls = 0;
    const originalFlushTelemetry = privateApi.flushTelemetry.bind(tracker);
    privateApi.flushTelemetry = (force?: boolean): void => {
      flushCalls += 1;
      originalFlushTelemetry(force);
    };

    tracker.handleMediaChange("/tmp/episode-2.mkv", "Episode 2");
    tracker.recordSubtitleLine("Hello immersion", 0, 1);
    tracker.destroy();

    assert.equal(flushCalls, 1);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test("monthly rollups are grouped by calendar month", async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    tracker = new ImmersionTrackerService({ dbPath });
    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      runRollupMaintenance: () => void;
    };

    const januaryStartedAtMs = Date.UTC(2026, 0, 31, 23, 59, 59, 0);
    const februaryStartedAtMs = Date.UTC(2026, 1, 1, 0, 0, 1, 0);

    privateApi.db.exec(`
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        created_at_ms,
        updated_at_ms
      ) VALUES (
        1,
        'local:/tmp/video.mkv',
        'Episode',
        1,
        0,
        ${januaryStartedAtMs},
        ${januaryStartedAtMs}
      )
    `);

    privateApi.db.exec(`
      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        status,
        created_at_ms,
        updated_at_ms,
        ended_at_ms
      ) VALUES (
        1,
        '11111111-1111-1111-1111-111111111111',
        1,
        ${januaryStartedAtMs},
        2,
        ${januaryStartedAtMs},
        ${januaryStartedAtMs},
        ${januaryStartedAtMs + 5000}
      )
    `);
    privateApi.db.exec(`
      INSERT INTO imm_session_telemetry (
        session_id,
        sample_ms,
        total_watched_ms,
        active_watched_ms,
        lines_seen,
        words_seen,
        tokens_seen,
        cards_mined,
        lookup_count,
        lookup_hits,
        pause_count,
        pause_ms,
        seek_forward_count,
        seek_backward_count,
        media_buffer_events
      ) VALUES (
        1,
        ${januaryStartedAtMs + 1000},
        5000,
        5000,
        1,
        2,
        2,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      )
    `);

    privateApi.db.exec(`
      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        status,
        created_at_ms,
        updated_at_ms,
        ended_at_ms
      ) VALUES (
        2,
        '22222222-2222-2222-2222-222222222222',
        1,
        ${februaryStartedAtMs},
        2,
        ${februaryStartedAtMs},
        ${februaryStartedAtMs},
        ${februaryStartedAtMs + 5000}
      )
    `);
    privateApi.db.exec(`
      INSERT INTO imm_session_telemetry (
        session_id,
        sample_ms,
        total_watched_ms,
        active_watched_ms,
        lines_seen,
        words_seen,
        tokens_seen,
        cards_mined,
        lookup_count,
        lookup_hits,
        pause_count,
        pause_ms,
        seek_forward_count,
        seek_backward_count,
        media_buffer_events
      ) VALUES (
        2,
        ${februaryStartedAtMs + 1000},
        4000,
        4000,
        2,
        3,
        3,
        1,
        1,
        1,
        0,
        0,
        0,
        0,
        0
      )
    `);

    privateApi.runRollupMaintenance();

    const rows = await tracker.getMonthlyRollups(10);
    const videoRows = rows.filter((row) => row.videoId === 1);

    assert.equal(videoRows.length, 2);
    assert.equal(videoRows[0].rollupDayOrMonth, 202602);
    assert.equal(videoRows[1].rollupDayOrMonth, 202601);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test("flushSingle reuses cached prepared statements", () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  let originalPrepare: DatabaseSync["prepare"] | null = null;

  try {
    tracker = new ImmersionTrackerService({ dbPath });
    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      flushSingle: (write: {
        kind: "telemetry" | "event";
        sessionId: number;
        sampleMs: number;
        eventType?: number;
        lineIndex?: number | null;
        segmentStartMs?: number | null;
        segmentEndMs?: number | null;
        wordsDelta?: number;
        cardsDelta?: number;
        payloadJson?: string | null;
        totalWatchedMs?: number;
        activeWatchedMs?: number;
        linesSeen?: number;
        wordsSeen?: number;
        tokensSeen?: number;
        cardsMined?: number;
        lookupCount?: number;
        lookupHits?: number;
        pauseCount?: number;
        pauseMs?: number;
        seekForwardCount?: number;
        seekBackwardCount?: number;
        mediaBufferEvents?: number;
      }) => void;
    };

    originalPrepare = privateApi.db.prepare;
    let prepareCalls = 0;
    privateApi.db.prepare = (...args: Parameters<DatabaseSync["prepare"]>) => {
      prepareCalls += 1;
      return originalPrepare!.apply(privateApi.db, args);
    };
    const preparedRestore = originalPrepare;

    privateApi.db.exec(`
      INSERT INTO imm_videos (
        video_id,
        video_key,
        canonical_title,
        source_type,
        duration_ms,
        created_at_ms,
        updated_at_ms
      ) VALUES (
        1,
        'local:/tmp/prepared.mkv',
        'Prepared',
        1,
        0,
        1000,
        1000
      )
    `);

    privateApi.db.exec(`
      INSERT INTO imm_sessions (
        session_id,
        session_uuid,
        video_id,
        started_at_ms,
        status,
        created_at_ms,
        updated_at_ms,
        ended_at_ms
      ) VALUES (
        1,
        '33333333-3333-3333-3333-333333333333',
        1,
        1000,
        2,
        1000,
        1000,
        2000
      )
    `);

    privateApi.flushSingle({
      kind: "telemetry",
      sessionId: 1,
      sampleMs: 1500,
      totalWatchedMs: 1000,
      activeWatchedMs: 1000,
      linesSeen: 1,
      wordsSeen: 2,
      tokensSeen: 2,
      cardsMined: 0,
      lookupCount: 0,
      lookupHits: 0,
      pauseCount: 0,
      pauseMs: 0,
      seekForwardCount: 0,
      seekBackwardCount: 0,
      mediaBufferEvents: 0,
    });

    privateApi.flushSingle({
      kind: "event",
      sessionId: 1,
      sampleMs: 1600,
      eventType: 1,
      lineIndex: 1,
      segmentStartMs: 0,
      segmentEndMs: 1000,
      wordsDelta: 2,
      cardsDelta: 0,
      payloadJson: '{"event":"subtitle-line"}',
    });

    privateApi.db.prepare = preparedRestore;

    assert.equal(prepareCalls, 0);
  } finally {
    if (tracker && originalPrepare) {
      const privateApi = tracker as unknown as { db: DatabaseSync };
      privateApi.db.prepare = originalPrepare;
    }
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});
