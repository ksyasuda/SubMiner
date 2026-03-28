import type { DatabaseSync } from './sqlite';
import { nowMs } from './time';
import { toDbMs } from './query-shared';

const ROLLUP_STATE_KEY = 'last_rollup_sample_ms';
const DAILY_MS = 86_400_000;
const ZERO_ID = 0;

interface RollupStateRow {
  state_value: number;
}

interface RollupGroupRow {
  rollup_day: number;
  rollup_month: number;
  video_id: number;
}

interface RollupTelemetryResult {
  maxSampleMs: number | null;
}

interface RawRetentionResult {
  deletedSessionEvents: number;
  deletedTelemetryRows: number;
  deletedEndedSessions: number;
}

export function toMonthKey(timestampMs: number): number {
  const epochDay = Math.floor(timestampMs / DAILY_MS);
  const z = epochDay + 719468;
  const era = Math.floor(z / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor(
    (doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365,
  );
  let year = yoe + era * 400;
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const month = mp + (mp < 10 ? 3 : -9);
  if (month <= 2) {
    year += 1;
  }
  return year * 100 + month;
}

export function pruneRawRetention(
  db: DatabaseSync,
  nowMs: number,
  policy: {
    eventsRetentionMs: number;
    telemetryRetentionMs: number;
    sessionsRetentionMs: number;
  },
): RawRetentionResult {
  const eventCutoff = nowMs - policy.eventsRetentionMs;
  const telemetryCutoff = nowMs - policy.telemetryRetentionMs;
  const sessionsCutoff = nowMs - policy.sessionsRetentionMs;

  const deletedSessionEvents = (
    db.prepare(`DELETE FROM imm_session_events WHERE ts_ms < ?`).run(toDbMs(eventCutoff)) as {
      changes: number;
    }
  ).changes;
  const deletedTelemetryRows = (
    db
      .prepare(`DELETE FROM imm_session_telemetry WHERE sample_ms < ?`)
      .run(toDbMs(telemetryCutoff)) as { changes: number }
  ).changes;
  const deletedEndedSessions = (
    db
      .prepare(`DELETE FROM imm_sessions WHERE ended_at_ms IS NOT NULL AND ended_at_ms < ?`)
      .run(toDbMs(sessionsCutoff)) as { changes: number }
  ).changes;

  return {
    deletedSessionEvents,
    deletedTelemetryRows,
    deletedEndedSessions,
  };
}

export function pruneRollupRetention(
  db: DatabaseSync,
  nowMs: number,
  policy: {
    dailyRollupRetentionMs: number;
    monthlyRollupRetentionMs: number;
  },
): { deletedDailyRows: number; deletedMonthlyRows: number } {
  const deletedDailyRows = Number.isFinite(policy.dailyRollupRetentionMs)
    ? (
        db
          .prepare(`DELETE FROM imm_daily_rollups WHERE rollup_day < ?`)
          .run(Math.floor((nowMs - policy.dailyRollupRetentionMs) / DAILY_MS)) as {
          changes: number;
        }
      ).changes
    : 0;
  const deletedMonthlyRows = Number.isFinite(policy.monthlyRollupRetentionMs)
    ? (
        db
          .prepare(`DELETE FROM imm_monthly_rollups WHERE rollup_month < ?`)
          .run(toMonthKey(nowMs - policy.monthlyRollupRetentionMs)) as {
          changes: number;
        }
      ).changes
    : 0;

  return {
    deletedDailyRows,
    deletedMonthlyRows,
  };
}

function getLastRollupSampleMs(db: DatabaseSync): number {
  const row = db
    .prepare(`SELECT state_value FROM imm_rollup_state WHERE state_key = ? LIMIT 1`)
    .get(ROLLUP_STATE_KEY) as unknown as RollupStateRow | null;
  return row ? Number(row.state_value) : ZERO_ID;
}

function setLastRollupSampleMs(db: DatabaseSync, sampleMs: number | bigint): void {
  db.prepare(
    `INSERT INTO imm_rollup_state (state_key, state_value)
       VALUES (?, ?)
       ON CONFLICT(state_key) DO UPDATE SET state_value = excluded.state_value`,
  ).run(ROLLUP_STATE_KEY, sampleMs);
}

function resetRollups(db: DatabaseSync): void {
  db.exec(`
    DELETE FROM imm_daily_rollups;
    DELETE FROM imm_monthly_rollups;
  `);
  setLastRollupSampleMs(db, ZERO_ID);
}

function upsertDailyRollupsForGroups(
  db: DatabaseSync,
  groups: Array<{ rollupDay: number; videoId: number }>,
  rollupNowMs: bigint,
): void {
  if (groups.length === 0) {
    return;
  }

  const upsertStmt = db.prepare(`
    INSERT INTO imm_daily_rollups (
      rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
      total_tokens_seen, total_cards, cards_per_hour,
      tokens_per_min, lookup_hit_rate, CREATED_DATE, LAST_UPDATE_DATE
    )
    SELECT
      CAST(julianday(s.started_at_ms / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) AS rollup_day,
      s.video_id AS video_id,
      COUNT(DISTINCT s.session_id) AS total_sessions,
      COALESCE(SUM(sm.max_active_ms), 0) / 60000.0 AS total_active_min,
      COALESCE(SUM(sm.max_lines), 0) AS total_lines_seen,
      COALESCE(SUM(sm.max_tokens), 0) AS total_tokens_seen,
      COALESCE(SUM(sm.max_cards), 0) AS total_cards,
      CASE
        WHEN COALESCE(SUM(sm.max_active_ms), 0) > 0
          THEN (COALESCE(SUM(sm.max_cards), 0) * 60.0) / (COALESCE(SUM(sm.max_active_ms), 0) / 60000.0)
        ELSE NULL
      END AS cards_per_hour,
      CASE
        WHEN COALESCE(SUM(sm.max_active_ms), 0) > 0
          THEN COALESCE(SUM(sm.max_tokens), 0) / (COALESCE(SUM(sm.max_active_ms), 0) / 60000.0)
        ELSE NULL
      END AS tokens_per_min,
      CASE
        WHEN COALESCE(SUM(sm.max_lookups), 0) > 0
          THEN CAST(COALESCE(SUM(sm.max_hits), 0) AS REAL) / CAST(SUM(sm.max_lookups) AS REAL)
        ELSE NULL
      END AS lookup_hit_rate,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM imm_sessions s
    JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.lines_seen) AS max_lines,
        MAX(t.tokens_seen) AS max_tokens,
        MAX(t.cards_mined) AS max_cards,
        MAX(t.lookup_count) AS max_lookups,
        MAX(t.lookup_hits) AS max_hits
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON s.session_id = sm.session_id
    WHERE CAST(julianday(s.started_at_ms / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) = ? AND s.video_id = ?
    GROUP BY rollup_day, s.video_id
    ON CONFLICT (rollup_day, video_id) DO UPDATE SET
      total_sessions = excluded.total_sessions,
      total_active_min = excluded.total_active_min,
      total_lines_seen = excluded.total_lines_seen,
      total_tokens_seen = excluded.total_tokens_seen,
      total_cards = excluded.total_cards,
      cards_per_hour = excluded.cards_per_hour,
      tokens_per_min = excluded.tokens_per_min,
      lookup_hit_rate = excluded.lookup_hit_rate,
      CREATED_DATE = COALESCE(imm_daily_rollups.CREATED_DATE, excluded.CREATED_DATE),
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `);

  for (const { rollupDay, videoId } of groups) {
    upsertStmt.run(rollupNowMs, rollupNowMs, rollupDay, videoId);
  }
}

function upsertMonthlyRollupsForGroups(
  db: DatabaseSync,
  groups: Array<{ rollupMonth: number; videoId: number }>,
  rollupNowMs: bigint,
): void {
  if (groups.length === 0) {
    return;
  }

  const upsertStmt = db.prepare(`
    INSERT INTO imm_monthly_rollups (
      rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
      total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
    )
    SELECT
      CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch', 'localtime') AS INTEGER) AS rollup_month,
      s.video_id AS video_id,
      COUNT(DISTINCT s.session_id) AS total_sessions,
      COALESCE(SUM(sm.max_active_ms), 0) / 60000.0 AS total_active_min,
      COALESCE(SUM(sm.max_lines), 0) AS total_lines_seen,
      COALESCE(SUM(sm.max_tokens), 0) AS total_tokens_seen,
      COALESCE(SUM(sm.max_cards), 0) AS total_cards,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM imm_sessions s
    JOIN (
      SELECT
        t.session_id,
        MAX(t.active_watched_ms) AS max_active_ms,
        MAX(t.lines_seen) AS max_lines,
        MAX(t.tokens_seen) AS max_tokens,
        MAX(t.cards_mined) AS max_cards
      FROM imm_session_telemetry t
      GROUP BY t.session_id
    ) sm ON s.session_id = sm.session_id
    WHERE CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch', 'localtime') AS INTEGER) = ? AND s.video_id = ?
    GROUP BY rollup_month, s.video_id
    ON CONFLICT (rollup_month, video_id) DO UPDATE SET
      total_sessions = excluded.total_sessions,
      total_active_min = excluded.total_active_min,
      total_lines_seen = excluded.total_lines_seen,
      total_tokens_seen = excluded.total_tokens_seen,
      total_cards = excluded.total_cards,
      CREATED_DATE = COALESCE(imm_monthly_rollups.CREATED_DATE, excluded.CREATED_DATE),
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `);

  for (const { rollupMonth, videoId } of groups) {
    upsertStmt.run(rollupNowMs, rollupNowMs, rollupMonth, videoId);
  }
}

function getAffectedRollupGroups(
  db: DatabaseSync,
  lastRollupSampleMs: number,
): Array<{ rollupDay: number; rollupMonth: number; videoId: number }> {
  return (
    db
      .prepare(
        `
          SELECT DISTINCT
            CAST(julianday(s.started_at_ms / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) AS rollup_day,
            CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch', 'localtime') AS INTEGER) AS rollup_month,
            s.video_id AS video_id
          FROM imm_session_telemetry t
          JOIN imm_sessions s
            ON s.session_id = t.session_id
          WHERE t.sample_ms > ?
        `,
      )
      .all(lastRollupSampleMs) as unknown as RollupGroupRow[]
  ).map((row) => ({
    rollupDay: row.rollup_day,
    rollupMonth: row.rollup_month,
    videoId: row.video_id,
  }));
}

function dedupeGroups<T extends { rollupDay?: number; rollupMonth?: number; videoId: number }>(
  groups: Array<T>,
): Array<T> {
  const seen = new Set<string>();
  const result: Array<T> = [];
  for (const group of groups) {
    const key = `${group.rollupDay ?? group.rollupMonth}-${group.videoId}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(group);
  }
  return result;
}

export function runRollupMaintenance(db: DatabaseSync, forceRebuild = false): void {
  if (forceRebuild) {
    db.exec('BEGIN IMMEDIATE');
    try {
      rebuildRollupsInTransaction(db);
      db.exec('COMMIT');
    } catch (error) {
      db.exec('ROLLBACK');
      throw error;
    }
    return;
  }

  const rollupNowMs = toDbMs(nowMs());
  const lastRollupSampleMs = getLastRollupSampleMs(db);

  const maxSampleRow = db
    .prepare('SELECT MAX(sample_ms) AS maxSampleMs FROM imm_session_telemetry')
    .get() as unknown as RollupTelemetryResult | null;
  if (!maxSampleRow?.maxSampleMs) {
    if (forceRebuild) {
      setLastRollupSampleMs(db, ZERO_ID);
    }
    return;
  }

  const affectedGroups = getAffectedRollupGroups(db, lastRollupSampleMs);
  if (!forceRebuild && affectedGroups.length === 0) {
    return;
  }

  const dailyGroups = dedupeGroups(
    affectedGroups.map((group) => ({
      rollupDay: group.rollupDay,
      videoId: group.videoId,
    })),
  );
  const monthlyGroups = dedupeGroups(
    affectedGroups.map((group) => ({
      rollupMonth: group.rollupMonth,
      videoId: group.videoId,
    })),
  );

  db.exec('BEGIN IMMEDIATE');
  try {
    upsertDailyRollupsForGroups(db, dailyGroups, rollupNowMs);
    upsertMonthlyRollupsForGroups(db, monthlyGroups, rollupNowMs);
    setLastRollupSampleMs(db, toDbMs(maxSampleRow.maxSampleMs ?? ZERO_ID));
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

export function rebuildRollupsInTransaction(db: DatabaseSync): void {
  const rollupNowMs = toDbMs(nowMs());
  const maxSampleRow = db
    .prepare('SELECT MAX(sample_ms) AS maxSampleMs FROM imm_session_telemetry')
    .get() as unknown as RollupTelemetryResult | null;

  resetRollups(db);
  if (!maxSampleRow?.maxSampleMs) {
    return;
  }

  const affectedGroups = getAffectedRollupGroups(db, ZERO_ID);
  if (affectedGroups.length === 0) {
    setLastRollupSampleMs(db, toDbMs(maxSampleRow.maxSampleMs ?? ZERO_ID));
    return;
  }

  const dailyGroups = dedupeGroups(
    affectedGroups.map((group) => ({
      rollupDay: group.rollupDay,
      videoId: group.videoId,
    })),
  );
  const monthlyGroups = dedupeGroups(
    affectedGroups.map((group) => ({
      rollupMonth: group.rollupMonth,
      videoId: group.videoId,
    })),
  );

  upsertDailyRollupsForGroups(db, dailyGroups, rollupNowMs);
  upsertMonthlyRollupsForGroups(db, monthlyGroups, rollupNowMs);
  setLastRollupSampleMs(db, toDbMs(maxSampleRow.maxSampleMs ?? ZERO_ID));
}

export function runOptimizeMaintenance(db: DatabaseSync): void {
  db.exec('PRAGMA optimize');
}
