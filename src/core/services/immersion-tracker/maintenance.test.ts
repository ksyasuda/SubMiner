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
    const nowMs = 90 * 86_400_000;
    const staleEndedAtMs = nowMs - 40 * 86_400_000;
    const keptEndedAtMs = nowMs - 5 * 86_400_000;

    db.exec(`
      INSERT INTO imm_videos (
        video_id, video_key, canonical_title, source_type, duration_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'local:/tmp/video.mkv', 'Video', 1, 0, ${nowMs}, ${nowMs}
      );
      INSERT INTO imm_sessions (
        session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES
        (1, 'session-1', 1, ${staleEndedAtMs - 1_000}, ${staleEndedAtMs}, 2, ${staleEndedAtMs}, ${staleEndedAtMs}),
        (2, 'session-2', 1, ${keptEndedAtMs - 1_000}, ${keptEndedAtMs}, 2, ${keptEndedAtMs}, ${keptEndedAtMs});
      INSERT INTO imm_session_telemetry (
        session_id, sample_ms, total_watched_ms, active_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES
        (1, ${nowMs - 2 * 86_400_000}, 0, 0, ${nowMs}, ${nowMs}),
        (2, ${nowMs - 12 * 60 * 60 * 1000}, 0, 0, ${nowMs}, ${nowMs});
    `);

    const result = pruneRawRetention(db, nowMs, {
      eventsRetentionMs: 7 * 86_400_000,
      telemetryRetentionMs: 1 * 86_400_000,
      sessionsRetentionMs: 30 * 86_400_000,
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

test('raw retention keeps rollups and rollup retention prunes them separately', () => {
  const dbPath = makeDbPath();
  const db = new Database(dbPath);

  try {
    ensureSchema(db);
    const nowMs = Date.UTC(2026, 2, 16, 12, 0, 0, 0);
    const oldDay = Math.floor((nowMs - 90 * 86_400_000) / 86_400_000);
    const oldMonth = toMonthKey(nowMs - 400 * 86_400_000);

    db.exec(`
      INSERT INTO imm_videos (
        video_id, video_key, canonical_title, source_type, duration_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'local:/tmp/video.mkv', 'Video', 1, 0, ${nowMs}, ${nowMs}
      );
      INSERT INTO imm_sessions (
        session_id, session_uuid, video_id, started_at_ms, ended_at_ms, status, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, 'session-1', 1, ${nowMs - 90 * 86_400_000}, ${nowMs - 90 * 86_400_000 + 1_000}, 2, ${nowMs}, ${nowMs}
      );
      INSERT INTO imm_session_telemetry (
        session_id, sample_ms, total_watched_ms, active_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        1, ${nowMs - 90 * 86_400_000}, 0, 0, ${nowMs}, ${nowMs}
      );
      INSERT INTO imm_daily_rollups (
        rollup_day, video_id, total_sessions, total_active_min, total_lines_seen, total_words_seen,
        total_tokens_seen, total_cards
      ) VALUES (
        ${oldDay}, 1, 1, 10, 1, 1, 1, 1
      );
      INSERT INTO imm_monthly_rollups (
        rollup_month, video_id, total_sessions, total_active_min, total_lines_seen, total_words_seen,
        total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
      ) VALUES (
        ${oldMonth}, 1, 1, 10, 1, 1, 1, 1, ${nowMs}, ${nowMs}
      );
    `);

    pruneRawRetention(db, nowMs, {
      eventsRetentionMs: 7 * 86_400_000,
      telemetryRetentionMs: 30 * 86_400_000,
      sessionsRetentionMs: 30 * 86_400_000,
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
      dailyRollupRetentionMs: 30 * 86_400_000,
      monthlyRollupRetentionMs: 365 * 86_400_000,
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
