import type { DatabaseSync } from 'node:sqlite';

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
): void {
  const eventCutoff = nowMs - policy.eventsRetentionMs;
  const telemetryCutoff = nowMs - policy.telemetryRetentionMs;
  const dailyCutoff = nowMs - policy.dailyRollupRetentionMs;
  const monthlyCutoff = nowMs - policy.monthlyRollupRetentionMs;
  const dayCutoff = Math.floor(dailyCutoff / 86_400_000);
  const monthCutoff = toMonthKey(monthlyCutoff);

  db.prepare(`DELETE FROM imm_session_events WHERE ts_ms < ?`).run(eventCutoff);
  db.prepare(`DELETE FROM imm_session_telemetry WHERE sample_ms < ?`).run(telemetryCutoff);
  db.prepare(`DELETE FROM imm_daily_rollups WHERE rollup_day < ?`).run(dayCutoff);
  db.prepare(`DELETE FROM imm_monthly_rollups WHERE rollup_month < ?`).run(monthCutoff);
  db.prepare(`DELETE FROM imm_sessions WHERE ended_at_ms IS NOT NULL AND ended_at_ms < ?`).run(
    telemetryCutoff,
  );
}

export function runRollupMaintenance(db: DatabaseSync): void {
  db.exec(`
    INSERT OR REPLACE INTO imm_daily_rollups (
      rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
      total_words_seen, total_tokens_seen, total_cards, cards_per_hour,
      words_per_min, lookup_hit_rate
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
      END AS lookup_hit_rate
    FROM imm_sessions s
    JOIN imm_session_telemetry t
      ON t.session_id = s.session_id
    GROUP BY rollup_day, s.video_id
  `);

  db.exec(`
    INSERT OR REPLACE INTO imm_monthly_rollups (
      rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
      total_words_seen, total_tokens_seen, total_cards
    )
    SELECT
      CAST(strftime('%Y%m', s.started_at_ms / 1000, 'unixepoch') AS INTEGER) AS rollup_month,
      s.video_id AS video_id,
      COUNT(DISTINCT s.session_id) AS total_sessions,
      COALESCE(SUM(t.active_watched_ms), 0) / 60000.0 AS total_active_min,
      COALESCE(SUM(t.lines_seen), 0) AS total_lines_seen,
      COALESCE(SUM(t.words_seen), 0) AS total_words_seen,
      COALESCE(SUM(t.tokens_seen), 0) AS total_tokens_seen,
      COALESCE(SUM(t.cards_mined), 0) AS total_cards
    FROM imm_sessions s
    JOIN imm_session_telemetry t
      ON t.session_id = s.session_id
    GROUP BY rollup_month, s.video_id
  `);
}
