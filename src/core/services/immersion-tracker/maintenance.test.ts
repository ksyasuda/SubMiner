import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Database } from './sqlite';
import {
  pruneRawRetention,
  pruneRollupRetention,
  runOptimizeMaintenance,
  toMonthKey,
} from './maintenance';
import { ensureSchema } from './storage';
import { toDbTimestamp } from './query-shared';

function makeDbPath(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-maintenance-test-'));
  return path.join(dir, 'tracker.db');
}

function cleanupDbPath(dbPath: string): void {
  try {
    fs.rmSync(path.dirname(dbPath), { recursive: true, force: true });
  } catch {
    // best effort
  }
}

test('pruneRawRetention uses session retention separately from telemetry retention', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const nowMs = 1_000_000_000;
    const staleEndedAtMs = nowMs - 400_000_000;
    const keptEndedAtMs = nowMs - 50_000_000;

    db.exec(`
      INSERT INTO imm_videos (
        video_id, video_key, canonical_title, source_type, duration_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'local:/tmp/video.mkv', 'Video', 1, 0, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
      INSERT INTO imm_sessions (
        session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES
        (1, 'session-1', 1, '${toDbTimestamp(staleEndedAtMs - 1_000)}', '${toDbTimestamp(staleEndedAtMs)}', 2, '${toDbTimestamp(staleEndedAtMs)}', '${toDbTimestamp(staleEndedAtMs)}'),
        (2, 'session-2', 1, '${toDbTimestamp(keptEndedAtMs - 1_000)}', '${toDbTimestamp(keptEndedAtMs)}', 2, '${toDbTimestamp(keptEndedAtMs)}', '${toDbTimestamp(keptEndedAtMs)}');
      INSERT INTO imm_session_telemetry (
        session_id, sample_ms, total_watched_ms, active_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES
        (1, '${toDbTimestamp(nowMs - 200_000_000)}', 0, 0, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'),
        (2, '${toDbTimestamp(nowMs - 10_000_000)}', 0, 0, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}');
    `);

    const result = pruneRawRetention(db, nowMs, {
      eventsRetentionMs: 120_000_000,
      telemetryRetentionMs: 80_000_000,
      sessionsRetentionMs: 300_000_000,
    });

    const remainingSessions = db
      .prepare('SELECT session_id FROM imm_sessions ORDER BY session_id')
      .all() as Array<{ session_id: number }>;
    const remainingTelemetry = db
      .prepare('SELECT session_id FROM imm_session_telemetry ORDER BY session_id')
      .all() as Array<{ session_id: number }>;

    assert.equal(result.deletedTelemetryRows, 1);
    assert.equal(result.deletedEndedSessions, 1);
    assert.deepEqual(
      remainingSessions.map((row) => row.session_id),
      [2],
    );
    assert.deepEqual(
      remainingTelemetry.map((row) => row.session_id),
      [2],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('pruneRawRetention skips disabled retention windows', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const nowMs = 1_000_000_000;

    db.exec(`
      INSERT INTO imm_videos (
        video_id, video_key, canonical_title, source_type, duration_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'local:/tmp/video.mkv', 'Video', 1, 0, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
      INSERT INTO imm_sessions (
        session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'session-1', 1, '${toDbTimestamp(nowMs - 1_000)}', '${toDbTimestamp(nowMs - 500)}', 2, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
      INSERT INTO imm_session_telemetry (
        session_id, sample_ms, total_watched_ms, active_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, '${toDbTimestamp(nowMs - 2_000)}', 0, 0, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
      INSERT INTO imm_session_events (
        session_id, event_type, ts_ms, payload_json, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 1, '${toDbTimestamp(nowMs - 3_000)}', '{}', '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
    `);

    const result = pruneRawRetention(db, nowMs, {
      eventsRetentionMs: Number.POSITIVE_INFINITY,
      telemetryRetentionMs: Number.POSITIVE_INFINITY,
      sessionsRetentionMs: Number.POSITIVE_INFINITY,
    });

    const remainingSessionEvents = db
      .prepare('SELECT COUNT(*) AS count FROM imm_session_events')
      .get() as { count: number };
    const remainingTelemetry = db
      .prepare('SELECT COUNT(*) AS count FROM imm_session_telemetry')
      .get() as { count: number };
    const remainingSessions = db
      .prepare('SELECT COUNT(*) AS count FROM imm_sessions')
      .get() as { count: number };

    assert.equal(result.deletedSessionEvents, 0);
    assert.equal(result.deletedTelemetryRows, 0);
    assert.equal(result.deletedEndedSessions, 0);
    assert.equal(remainingSessionEvents.count, 1);
    assert.equal(remainingTelemetry.count, 1);
    assert.equal(remainingSessions.count, 1);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('toMonthKey floors negative timestamps into the prior UTC month', () => {
  assert.equal(toMonthKey(-1), 196912);
  assert.equal(toMonthKey(-86_400_000), 196912);
  assert.equal(toMonthKey(0), 197001);
});

test('raw retention keeps rollups and rollup retention prunes them separately', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const nowMs = 1_000_000_000;
    const oldDay = Math.floor((nowMs - 200_000_000) / 86_400_000);
    const oldMonth = 196912;

    db.exec(`
      INSERT INTO imm_videos (
        video_id, video_key, canonical_title, source_type, duration_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'local:/tmp/video.mkv', 'Video', 1, 0, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
      INSERT INTO imm_sessions (
        session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'session-1', 1, '${toDbTimestamp(nowMs - 200_000_000)}', '${toDbTimestamp(nowMs - 199_999_000)}', 2, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
      INSERT INTO imm_session_telemetry (
        session_id, sample_ms, total_watched_ms, active_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, '${toDbTimestamp(nowMs - 200_000_000)}', 0, 0, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
      INSERT INTO imm_daily_rollups (
        rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
        total_tokens_seen, total_cards
      ) VALUES (
        ${oldDay}, 1, 1, 10, 1, 1, 1
      );
      INSERT INTO imm_monthly_rollups (
        rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
        total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        ${oldMonth}, 1, 1, 10, 1, 1, 1, '${toDbTimestamp(nowMs)}', '${toDbTimestamp(nowMs)}'
      );
    `);

    pruneRawRetention(db, nowMs, {
      eventsRetentionMs: 120_000_000,
      telemetryRetentionMs: 120_000_000,
      sessionsRetentionMs: 120_000_000,
    });

    const rollupsAfterRawPrune = db
      .prepare('SELECT COUNT(*) AS total FROM imm_daily_rollups')
      .get() as { total: number } | null;
    const monthlyAfterRawPrune = db
      .prepare('SELECT COUNT(*) AS total FROM imm_monthly_rollups')
      .get() as { total: number } | null;

    assert.equal(rollupsAfterRawPrune?.total, 1);
    assert.equal(monthlyAfterRawPrune?.total, 1);

    const rollupPrune = pruneRollupRetention(db, nowMs, {
      dailyRollupRetentionMs: 120_000_000,
      monthlyRollupRetentionMs: 1,
    });

    const rollupsAfterRollupPrune = db
      .prepare('SELECT COUNT(*) AS total FROM imm_daily_rollups')
      .get() as { total: number } | null;
    const monthlyAfterRollupPrune = db
      .prepare('SELECT COUNT(*) AS total FROM imm_monthly_rollups')
      .get() as { total: number } | null;

    assert.equal(rollupPrune.deletedDailyRows, 1);
    assert.equal(rollupPrune.deletedMonthlyRows, 1);
    assert.equal(rollupsAfterRollupPrune?.total, 0);
    assert.equal(monthlyAfterRollupPrune?.total, 0);
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('ensureSchema adds sample_ms index for telemetry rollup scans', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const indexes = db.prepare("PRAGMA index_list('imm_session_telemetry')").all() as Array<{
      name: string;
    }>;
    const hasSampleMsIndex = indexes.some((row) => row.name === 'idx_telemetry_sample_ms');
    assert.equal(hasSampleMsIndex, true);

    const indexColumns = db.prepare("PRAGMA index_info('idx_telemetry_sample_ms')").all() as Array<{
      name: string;
    }>;
    assert.deepEqual(
      indexColumns.map((column) => column.name),
      ['sample_ms'],
    );
  } finally {
    db.close();
    cleanupDbPath(dbPath);
  }
});

test('runOptimizeMaintenance executes PRAGMA optimize', () => {
  const executedSql: string[] = [];
  const db = {
    exec(source: string) {
      executedSql.push(source);
      return this;
    },
  } as unknown as Parameters<typeof runOptimizeMaintenance>[0];

  runOptimizeMaintenance(db);

  assert.deepEqual(executedSql, ['PRAGMA optimize']);
});
