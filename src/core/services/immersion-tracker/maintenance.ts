import type { DatabaseSync } from 'node:sqlite';

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

interface RetentionResult {
  deletedSessionEvents: number;
  deletedTelemetryRows: number;
  deletedDailyRows: number;
  deletedMonthlyRows: number;
  deletedEndedSessions: number;
}

export function toMonthKey(timestampMs: number): number {
  const monthDate = new Date(timestampMs);
  return monthDate.getUTCFullYear() * 100 + monthDate.getUTCMonth() + 1;
}

export function pruneRetention(
  db: DatabaseSync,
  nowMs: number,
  policy: {
    eventsRetentionMs: number;
    telemetryRetentionMs: number;
    dailyRollupRetentionMs: number;
    monthlyRollupRetentionMs: number;
  },
): RetentionResult {
  const eventCutoff = nowMs - policy.eventsRetentionMs;
  const telemetryCutoff = nowMs - policy.telemetryRetentionMs;
  const dayCutoff = nowMs - policy.dailyRollupRetentionMs;
  const monthCutoff = nowMs - policy.monthlyRollupRetentionMs;

  const deletedSessionEvents = (db
    .prepare(`DELETE FROM imm_session_events WHERE ts_ms < ?`)
    .run(eventCutoff) as { changes: number }).changes;
  const deletedTelemetryRows = (db
    .prepare(`DELETE FROM imm_session_telemetry WHERE sample_ms < ?`)
    .run(telemetryCutoff) as { changes: number }).changes;
  const deletedDailyRows = (db
    .prepare(`DELETE FROM imm_daily_rollups WHERE rollup_day < ?`)
    .run(Math.floor(dayCutoff / DAILY_MS)) as { changes: number }).changes;
  const deletedMonthlyRows = (db
    .prepare(`DELETE FROM imm_monthly_rollups WHERE rollup_month < ?`)
    .run(toMonthKey(monthCutoff)) as { changes: number }).changes;
  const deletedEndedSessions = (db
    .prepare(
      `DELETE FROM imm_sessions WHERE ended_at_ms IS NOT NULL AND ended_at_ms < ?`,
    )
    .run(telemetryCutoff) as { changes: number }).changes;

  return {
    deletedSessionEvents,
    deletedTelemetryRows,
    deletedDailyRows,
    deletedMonthlyRows,
    deletedEndedSessions,
  };
}

function getLastRollupSampleMs(db: DatabaseSync): number {
  const row = db
    .prepare(`SELECT state_value FROM imm_rollup_state WHERE state_key = ? LIMIT 1`)
    .get(ROLLUP_STATE_KEY) as unknown as RollupStateRow | null;
  return row ? Number(row.state_value) : ZERO_ID;
}

function setLastRollupSampleMs(db: DatabaseSync, sampleMs: number): void {
  db.prepare(
    `INSERT INTO imm_rollup_state (state_key, state_value)
       VALUES (?, ?)
       ON CONFLICT(state_key) DO UPDATE SET state_value = excluded.state_value`,
  ).run(ROLLUP_STATE_KEY, sampleMs);
}

function upsertDailyRollupsForGroups(
  db: DatabaseSync,
  groups: Array<{ rollupDay: number; videoId: number }>,
  rollupNowMs: number,
): void {
  if (groups.length === 0) {
    return;
  }

  const deleteStmt = db.prepare(`DELETE FROM imm_daily_rollups WHERE rollup_day = ? AND video_id = ?`);
  const upsertStmt = db.prepare(`
    INSERT INTO imm_daily_rollups (
      rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
      total_words_seen, total_tokens_seen, total_cards, cards_per_hour,
      words_per_min, lookup_hit_rate, CREATED_DATE, LAST_UPDATE_DATE
    )
    SELECT
      CAST(s.started_at_ms / 86400000 AS INTEGER) AS rollup_day,
      s.video_id AS video_id,
      COUNT(DISTINCT s.session_id) AS total_sessions,
      COALESCE(SUM(t.active_watched_ms), 0) / 60000.0 AS total_active_min,
      COALESCE(SUM(t.lines_seen), 0) AS total_lines_seen,
      COALESCE(SUM(t.words_seen), 0) AS total_words_seen,
      COALESCE(SUM(t.tokens_seen), 0) AS total_tokens_seen,
      COALESCE(SUM(t.cards_mined), 0) AS total_cards,
      CASE
        WHEN COALESCE(SUM(t.active_watched_ms), 0) > 0
          THEN (COALESCE(SUM(t.cards_mined), 0) * 60.0) / (COALESCE(SUM(t.active_watched_ms), 0) / 60000.0)
        ELSE NULL
      END AS cards_per_hour,
      CASE
        WHEN COALESCE(SUM(t.active_watched_ms), 0) > 0
          THEN COALESCE(SUM(t.words_seen), 0) / (COALESCE(SUM(t.active_watched_ms), 0) / 60000.0)
        ELSE NULL
      END AS words_per_min,
      CASE
        WHEN COALESCE(SUM(t.lookup_count), 0) > 0
          THEN CAST(COALESCE(SUM(t.lookup_hits), 0) AS REAL) / CAST(SUM(t.lookup_count) AS REAL)
        ELSE NULL
      END AS lookup_hit_rate,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM imm_sessions s
    JOIN imm_session_telemetry t
      ON t.session_id = s.session_id
    WHERE CAST(s.started_at_ms / 86400000 AS INTEGER) = ? AND s.video_id = ?
    GROUP BY rollup_day, s.video_id
    ON CONFLICT (rollup_day, video_id) DO UPDATE SET
      total_sessions = excluded.total_sessions,
      total_active_min = excluded.total_active_min,
      total_lines_seen = excluded.total_lines_seen,
      total_words_seen = excluded.total_words_seen,
      total_tokens_seen = excluded.total_tokens_seen,
      total_cards = excluded.total_cards,
      cards_per_hour = excluded.cards_per_hour,
      words_per_min = excluded.words_per_min,
      lookup_hit_rate = excluded.lookup_hit_rate,
      CREATED_DATE = COALESCE(imm_daily_rollups.CREATED_DATE, excluded.CREATED_DATE),
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `);

  for (const { rollupDay, videoId } of groups) {
    deleteStmt.run(rollupDay, videoId);
    upsertStmt.run(rollupNowMs, rollupNowMs, rollupDay, videoId);
  }
}

function upsertMonthlyRollupsForGroups(
  db: DatabaseSync,
  groups: Array<{ rollupMonth: number; videoId: number }>,
  rollupNowMs: number,
): void {
  if (groups.length === 0) {
    return;
  }

  const deleteStmt = db.prepare(
    `DELETE FROM imm_monthly_rollups WHERE rollup_month = ? AND video_id = ?`,
  );
  const upsertStmt = db.prepare(`
    INSERT INTO imm_monthly_rollups (
      rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
      total_words_seen, total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
    )
    SELECT
      CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch') AS INTEGER) AS rollup_month,
      s.video_id AS video_id,
      COUNT(DISTINCT s.session_id) AS total_sessions,
      COALESCE(SUM(t.active_watched_ms), 0) / 60000.0 AS total_active_min,
      COALESCE(SUM(t.lines_seen), 0) AS total_lines_seen,
      COALESCE(SUM(t.words_seen), 0) AS total_words_seen,
      COALESCE(SUM(t.tokens_seen), 0) AS total_tokens_seen,
      COALESCE(SUM(t.cards_mined), 0) AS total_cards,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM imm_sessions s
    JOIN imm_session_telemetry t
      ON t.session_id = s.session_id
    WHERE CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch') AS INTEGER) = ? AND s.video_id = ?
    GROUP BY rollup_month, s.video_id
    ON CONFLICT (rollup_month, video_id) DO UPDATE SET
      total_sessions = excluded.total_sessions,
      total_active_min = excluded.total_active_min,
      total_lines_seen = excluded.total_lines_seen,
      total_words_seen = excluded.total_words_seen,
      total_tokens_seen = excluded.total_tokens_seen,
      total_cards = excluded.total_cards,
      CREATED_DATE = COALESCE(imm_monthly_rollups.CREATED_DATE, excluded.CREATED_DATE),
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `);

  for (const { rollupMonth, videoId } of groups) {
    deleteStmt.run(rollupMonth, videoId);
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
            CAST(s.started_at_ms / 86400000 AS INTEGER) AS rollup_day,
            CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch') AS INTEGER) AS rollup_month,
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
  const rollupNowMs = Date.now();
  const lastRollupSampleMs = forceRebuild ? ZERO_ID : getLastRollupSampleMs(db);

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

  upsertDailyRollupsForGroups(db, dailyGroups, rollupNowMs);
  upsertMonthlyRollupsForGroups(db, monthlyGroups, rollupNowMs);
  setLastRollupSampleMs(db, Number(maxSampleRow.maxSampleMs));
}
