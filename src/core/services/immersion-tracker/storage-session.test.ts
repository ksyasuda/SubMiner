import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import type { DatabaseSync as NodeDatabaseSync } from 'node:sqlite';
import { finalizeSessionRecord, startSessionRecord } from './session';
import {
  createTrackerPreparedStatements,
  ensureSchema,
  executeQueuedWrite,
  getOrCreateVideoRecord,
} from './storage';
import { EVENT_SUBTITLE_LINE, SESSION_STATUS_ENDED, SOURCE_TYPE_LOCAL } from './types';

type DatabaseSyncCtor = typeof NodeDatabaseSync;
const DatabaseSync: DatabaseSyncCtor | null = (() => {
  try {
    return (require('node:sqlite') as { DatabaseSync?: DatabaseSyncCtor }).DatabaseSync ?? null;
  } catch {
    return null;
  }
})();
const testIfSqlite = DatabaseSync ? test : test.skip;

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-imm-storage-session-'));
  return path.join(dir, 'immersion.sqlite');
}

function cleanupDbPath(dbPath: string): void {
  const dir = path.dirname(dbPath);
  if (fs.existsSync(dir)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

testIfSqlite('ensureSchema creates immersion core tables', () => {
  const dbPath = makeDbPath();
  const db = new DatabaseSync!(dbPath);

  try {
    ensureSchema(db);
    const rows = db
      .prepare(
        `SELECT name FROM sqlite_master WHERE type = 'table' AND name LIKE 'imm_%' ORDER BY name`,
      )
      .all() as Array<{ name: string }>;
    const tableNames = new Set(rows.map((row) => row.name));

    assert.ok(tableNames.has('imm_videos'));
    assert.ok(tableNames.has('imm_sessions'));
    assert.ok(tableNames.has('imm_session_telemetry'));
    assert.ok(tableNames.has('imm_session_events'));
    assert.ok(tableNames.has('imm_daily_rollups'));
    assert.ok(tableNames.has('imm_monthly_rollups'));
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

testIfSqlite('start/finalize session updates ended_at and status', () => {
  const dbPath = makeDbPath();
  const db = new DatabaseSync!(dbPath);

  try {
    ensureSchema(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/slice-a.mkv', {
      canonicalTitle: 'Slice A Episode',
      sourcePath: '/tmp/slice-a.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const startedAtMs = 1_234_567_000;
    const endedAtMs = startedAtMs + 8_500;
    const { sessionId, state } = startSessionRecord(db, videoId, startedAtMs);

    finalizeSessionRecord(db, state, endedAtMs);

    const row = db
      .prepare('SELECT ended_at_ms, status FROM imm_sessions WHERE session_id = ?')
      .get(sessionId) as {
      ended_at_ms: number | null;
      status: number;
    } | null;

    assert.ok(row);
    assert.equal(row?.ended_at_ms, endedAtMs);
    assert.equal(row?.status, SESSION_STATUS_ENDED);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

testIfSqlite('executeQueuedWrite inserts event and telemetry rows', () => {
  const dbPath = makeDbPath();
  const db = new DatabaseSync!(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);
    const videoId = getOrCreateVideoRecord(db, 'local:/tmp/slice-a-events.mkv', {
      canonicalTitle: 'Slice A Events',
      sourcePath: '/tmp/slice-a-events.mkv',
      sourceUrl: null,
      sourceType: SOURCE_TYPE_LOCAL,
    });
    const { sessionId } = startSessionRecord(db, videoId, 5_000);

    executeQueuedWrite(
      {
        kind: 'telemetry',
        sessionId,
        sampleMs: 6_000,
        totalWatchedMs: 1_000,
        activeWatchedMs: 900,
        linesSeen: 3,
        wordsSeen: 6,
        tokensSeen: 6,
        cardsMined: 1,
        lookupCount: 2,
        lookupHits: 1,
        pauseCount: 1,
        pauseMs: 50,
        seekForwardCount: 0,
        seekBackwardCount: 0,
        mediaBufferEvents: 0,
      },
      stmts,
    );
    executeQueuedWrite(
      {
        kind: 'event',
        sessionId,
        sampleMs: 6_100,
        eventType: EVENT_SUBTITLE_LINE,
        lineIndex: 1,
        segmentStartMs: 0,
        segmentEndMs: 800,
        wordsDelta: 2,
        cardsDelta: 0,
        payloadJson: '{"event":"subtitle-line"}',
      },
      stmts,
    );

    const telemetryCount = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_telemetry WHERE session_id = ?')
      .get(sessionId) as { total: number };
    const eventCount = db
      .prepare('SELECT COUNT(*) AS total FROM imm_session_events WHERE session_id = ?')
      .get(sessionId) as { total: number };

    assert.equal(telemetryCount.total, 1);
    assert.equal(eventCount.total, 1);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
