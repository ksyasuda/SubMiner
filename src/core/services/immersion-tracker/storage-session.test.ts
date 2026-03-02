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
    assert.ok(tableNames.has('imm_words'));
    assert.ok(tableNames.has('imm_kanji'));
    assert.ok(tableNames.has('imm_rollup_state'));

    const rollupStateRow = db
      .prepare('SELECT state_value FROM imm_rollup_state WHERE state_key = ?')
      .get('last_rollup_sample_ms') as {
      state_value: number;
    } | null;
    assert.ok(rollupStateRow);
    assert.equal(rollupStateRow?.state_value, 0);
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

testIfSqlite('executeQueuedWrite inserts and upserts word and kanji rows', () => {
  const dbPath = makeDbPath();
  const db = new DatabaseSync!(dbPath);

  try {
    ensureSchema(db);
    const stmts = createTrackerPreparedStatements(db);

    stmts.wordUpsertStmt.run('猫', '猫', '', 10.0, 10.0);
    stmts.wordUpsertStmt.run('猫', '猫', '', 5.0, 15.0);
    stmts.kanjiUpsertStmt.run('日', 9.0, 9.0);
    stmts.kanjiUpsertStmt.run('日', 8.0, 11.0);

    const wordRow = db
      .prepare(
        'SELECT headword, frequency, first_seen, last_seen FROM imm_words WHERE headword = ?',
      )
      .get('猫') as {
      headword: string;
      frequency: number;
      first_seen: number;
      last_seen: number;
    } | null;
    const kanjiRow = db
      .prepare('SELECT kanji, frequency, first_seen, last_seen FROM imm_kanji WHERE kanji = ?')
      .get('日') as {
      kanji: string;
      frequency: number;
      first_seen: number;
      last_seen: number;
    } | null;

    assert.ok(wordRow);
    assert.ok(kanjiRow);
    assert.equal(wordRow?.frequency, 2);
    assert.equal(kanjiRow?.frequency, 2);
    assert.equal(wordRow?.first_seen, 5);
    assert.equal(wordRow?.last_seen, 15);
    assert.equal(kanjiRow?.first_seen, 8);
    assert.equal(kanjiRow?.last_seen, 11);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});
