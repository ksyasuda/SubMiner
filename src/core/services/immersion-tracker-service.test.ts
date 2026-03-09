import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { toMonthKey } from './immersion-tracker/maintenance';
import { enqueueWrite } from './immersion-tracker/queue';
import { Database, type DatabaseSync } from './immersion-tracker/sqlite';
import {
  deriveCanonicalTitle,
  normalizeText,
  resolveBoundedInt,
} from './immersion-tracker/reducer';
import type { QueuedWrite } from './immersion-tracker/types';

type ImmersionTrackerService = import('./immersion-tracker-service').ImmersionTrackerService;
type ImmersionTrackerServiceCtor =
  typeof import('./immersion-tracker-service').ImmersionTrackerService;

let trackerCtor: ImmersionTrackerServiceCtor | null = null;

async function loadTrackerCtor(): Promise<ImmersionTrackerServiceCtor> {
  if (trackerCtor) return trackerCtor;
  const mod = await import('./immersion-tracker-service');
  trackerCtor = mod.ImmersionTrackerService;
  return trackerCtor;
}

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-immersion-test-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) {
    return;
  }

  const bunRuntime = globalThis as typeof globalThis & {
    Bun?: {
      gc?: (force?: boolean) => void;
    };
  };
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
      return;
    } catch (error) {
      const err = error as NodeJS.ErrnoException;
      if (process.platform !== 'win32' || err.code !== 'EBUSY') {
        throw error;
      }
      bunRuntime.Bun?.gc?.(true);
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 25);
    }
  }

  // libsql keeps Windows file handles alive after close when prepared statements were used.
}

test('seam: resolveBoundedInt keeps fallback for invalid values', () => {
  assert.equal(resolveBoundedInt(undefined, 25, 1, 100), 25);
  assert.equal(resolveBoundedInt(0, 25, 1, 100), 25);
  assert.equal(resolveBoundedInt(101, 25, 1, 100), 25);
  assert.equal(resolveBoundedInt(44.8, 25, 1, 100), 44);
});

test('seam: reducer title normalization covers local and remote paths', () => {
  assert.equal(normalizeText('  hello\n  world  '), 'hello world');
  assert.equal(deriveCanonicalTitle('/tmp/Episode 01.mkv'), 'Episode 01');
  assert.equal(
    deriveCanonicalTitle('https://cdn.example.com/show/%E7%AC%AC1%E8%A9%B1.mp4'),
    '\u7b2c1\u8a71',
  );
});

test('seam: enqueueWrite drops oldest entries once capacity is exceeded', () => {
  const queue: QueuedWrite[] = [
    { kind: 'event', sessionId: 1, eventType: 1, sampleMs: 1000 },
    { kind: 'event', sessionId: 1, eventType: 2, sampleMs: 1001 },
  ];
  const incoming: QueuedWrite = { kind: 'event', sessionId: 1, eventType: 3, sampleMs: 1002 };

  const result = enqueueWrite(queue, incoming, 2);
  assert.equal(result.dropped, 1);
  assert.equal(queue.length, 2);
  assert.equal((queue[0] as Extract<QueuedWrite, { kind: 'event' }>).eventType, 2);
  assert.equal((queue[1] as Extract<QueuedWrite, { kind: 'event' }>).eventType, 3);
});

test('seam: toMonthKey uses UTC calendar month', () => {
  assert.equal(toMonthKey(Date.UTC(2026, 0, 31, 23, 59, 59, 999)), 202601);
  assert.equal(toMonthKey(Date.UTC(2026, 1, 1, 0, 0, 0, 0)), 202602);
});

test('startSession generates UUID-like session identifiers', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    tracker.handleMediaChange('/tmp/episode.mkv', 'Episode');

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const db = new Database(dbPath);
    const row = db.prepare('SELECT session_uuid FROM imm_sessions LIMIT 1').get() as {
      session_uuid: string;
    } | null;
    db.close();

    assert.equal(typeof row?.session_uuid, 'string');
    assert.equal(row?.session_uuid?.startsWith('session-'), false);
    assert.ok(/^[0-9a-fA-F-]{36}$/.test(row?.session_uuid || ''));
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('destroy finalizes active session and persists final telemetry', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/episode-2.mkv', 'Episode 2');
    tracker.recordSubtitleLine('Hello immersion', 0, 1);
    tracker.destroy();

    const db = new Database(dbPath);
    const sessionRow = db.prepare('SELECT ended_at_ms FROM imm_sessions LIMIT 1').get() as {
      ended_at_ms: number | null;
    } | null;
    const telemetryCountRow = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_telemetry')
      .get() as { total: number };
    db.close();

    assert.ok(sessionRow);
    assert.ok(Number(sessionRow?.ended_at_ms ?? 0) > 0);
    assert.ok(Number(telemetryCountRow.total) >= 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('persists and retrieves minimum immersion tracking fields', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });

    tracker.handleMediaChange('/tmp/episode-3.mkv', 'Episode 3');
    tracker.recordSubtitleLine('alpha beta', 0, 1.2);
    tracker.recordCardsMined(2);
    tracker.recordLookup(true);
    tracker.recordPlaybackPosition(12.5);

    const privateApi = tracker as unknown as {
      flushTelemetry: (force?: boolean) => void;
      flushNow: () => void;
    };
    privateApi.flushTelemetry(true);
    privateApi.flushNow();

    const summaries = await tracker.getSessionSummaries(10);
    assert.ok(summaries.length >= 1);
    assert.ok(summaries[0]!.linesSeen >= 1);
    assert.ok(summaries[0]!.cardsMined >= 2);

    tracker.destroy();

    const db = new Database(dbPath);
    const videoRow = db
      .prepare('SELECT canonical_title, source_path, duration_ms FROM imm_videos LIMIT 1')
      .get() as {
      canonical_title: string;
      source_path: string | null;
      duration_ms: number;
    } | null;
    const telemetryRow = db
      .prepare(
        `SELECT lines_seen, words_seen, tokens_seen, cards_mined
         FROM imm_session_telemetry
         ORDER BY sample_ms DESC, telemetry_id DESC
         LIMIT 1`,
      )
      .get() as {
      lines_seen: number;
      words_seen: number;
      tokens_seen: number;
      cards_mined: number;
    } | null;
    db.close();

    assert.ok(videoRow);
    assert.equal(videoRow?.canonical_title, 'Episode 3');
    assert.equal(videoRow?.source_path, '/tmp/episode-3.mkv');
    assert.ok(Number(videoRow?.duration_ms ?? -1) >= 0);

    assert.ok(telemetryRow);
    assert.ok(Number(telemetryRow?.lines_seen ?? 0) >= 1);
    assert.ok(Number(telemetryRow?.words_seen ?? 0) >= 2);
    assert.ok(Number(telemetryRow?.tokens_seen ?? 0) >= 2);
    assert.ok(Number(telemetryRow?.cards_mined ?? 0) >= 2);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('applies configurable queue, flush, and retention policy', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({
      dbPath,
      policy: {
        batchSize: 10,
        flushIntervalMs: 250,
        queueCap: 1500,
        payloadCapBytes: 512,
        maintenanceIntervalMs: 2 * 60 * 60 * 1000,
        retention: {
          eventsDays: 14,
          telemetryDays: 45,
          dailyRollupsDays: 730,
          monthlyRollupsDays: 3650,
          vacuumIntervalDays: 14,
        },
      },
    });

    const privateApi = tracker as unknown as {
      batchSize: number;
      flushIntervalMs: number;
      queueCap: number;
      maxPayloadBytes: number;
      maintenanceIntervalMs: number;
      eventsRetentionMs: number;
      telemetryRetentionMs: number;
      dailyRollupRetentionMs: number;
      monthlyRollupRetentionMs: number;
      vacuumIntervalMs: number;
    };

    assert.equal(privateApi.batchSize, 10);
    assert.equal(privateApi.flushIntervalMs, 250);
    assert.equal(privateApi.queueCap, 1500);
    assert.equal(privateApi.maxPayloadBytes, 512);
    assert.equal(privateApi.maintenanceIntervalMs, 7_200_000);
    assert.equal(privateApi.eventsRetentionMs, 14 * 86_400_000);
    assert.equal(privateApi.telemetryRetentionMs, 45 * 86_400_000);
    assert.equal(privateApi.dailyRollupRetentionMs, 730 * 86_400_000);
    assert.equal(privateApi.monthlyRollupRetentionMs, 3650 * 86_400_000);
    assert.equal(privateApi.vacuumIntervalMs, 14 * 86_400_000);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('monthly rollups are grouped by calendar month', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
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
        CREATED_DATE,
        LAST_UPDATE_DATE
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
        CREATED_DATE,
        LAST_UPDATE_DATE,
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
        CREATED_DATE,
        LAST_UPDATE_DATE,
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
    assert.equal(videoRows[0]!.rollupDayOrMonth, 202602);
    assert.equal(videoRows[1]!.rollupDayOrMonth, 202601);
  } finally {
    tracker?.destroy();
    cleanupDbPath(dbPath);
  }
});

test('flushSingle reuses cached prepared statements', async () => {
  const dbPath = makeDbPath();
  let tracker: ImmersionTrackerService | null = null;
  let originalPrepare: DatabaseSync['prepare'] | null = null;

  try {
    const Ctor = await loadTrackerCtor();
    tracker = new Ctor({ dbPath });
    const privateApi = tracker as unknown as {
      db: DatabaseSync;
      flushSingle: (write: {
        kind: 'telemetry' | 'event';
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
    privateApi.db.prepare = (...args: Parameters<DatabaseSync['prepare']>) => {
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
        CREATED_DATE,
        LAST_UPDATE_DATE
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
        CREATED_DATE,
        LAST_UPDATE_DATE,
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
      kind: 'telemetry',
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
      kind: 'event',
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
