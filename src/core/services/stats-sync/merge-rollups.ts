import { selectAll, type SqlRow, type SyncDb } from './libsql-driver';
import { nowDbTimestamp, tableExists, type SyncMergeSummary } from './shared';

const LOCAL_DAY_EXPR = `CAST(julianday(CAST(started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER)`;
const LOCAL_MONTH_EXPR = `CAST(strftime('%Y%m', CAST(started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') AS INTEGER)`;

// Ported from upsertDailyRollupsForGroups / upsertMonthlyRollupsForGroups in
// src/core/services/immersion-tracker/maintenance.ts — must stay in sync.
const DAILY_ROLLUP_UPSERT = `
  WITH matching_sessions AS (
    SELECT * FROM imm_sessions
    WHERE ${LOCAL_DAY_EXPR} = ? AND video_id = ?
  ),
  session_metrics AS (
    SELECT
      t.session_id,
      MAX(t.active_watched_ms) AS max_active_ms,
      MAX(t.lines_seen) AS max_lines,
      MAX(t.tokens_seen) AS max_tokens,
      MAX(t.cards_mined) AS max_cards,
      MAX(t.lookup_count) AS max_lookups,
      MAX(t.lookup_hits) AS max_hits
    FROM imm_session_telemetry t
    JOIN matching_sessions s ON s.session_id = t.session_id
    GROUP BY t.session_id
  )
  INSERT INTO imm_daily_rollups (
    rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
    total_tokens_seen, total_cards, cards_per_hour, tokens_per_min, lookup_hit_rate,
    CREATED_DATE, LAST_UPDATE_DATE
  )
  SELECT
    ${LOCAL_DAY_EXPR.replace('started_at_ms', 's.started_at_ms')} AS rollup_day,
    s.video_id AS video_id,
    COUNT(DISTINCT s.session_id) AS total_sessions,
    COALESCE(SUM(COALESCE(sm.max_active_ms, s.active_watched_ms)), 0) / 60000.0 AS total_active_min,
    COALESCE(SUM(COALESCE(sm.max_lines, s.lines_seen)), 0) AS total_lines_seen,
    COALESCE(SUM(COALESCE(sm.max_tokens, s.tokens_seen)), 0) AS total_tokens_seen,
    COALESCE(SUM(COALESCE(sm.max_cards, s.cards_mined)), 0) AS total_cards,
    CASE
      WHEN COALESCE(SUM(COALESCE(sm.max_active_ms, s.active_watched_ms)), 0) > 0
        THEN (COALESCE(SUM(COALESCE(sm.max_cards, s.cards_mined)), 0) * 60.0)
          / (COALESCE(SUM(COALESCE(sm.max_active_ms, s.active_watched_ms)), 0) / 60000.0)
      ELSE NULL
    END AS cards_per_hour,
    CASE
      WHEN COALESCE(SUM(COALESCE(sm.max_active_ms, s.active_watched_ms)), 0) > 0
        THEN COALESCE(SUM(COALESCE(sm.max_tokens, s.tokens_seen)), 0)
          / (COALESCE(SUM(COALESCE(sm.max_active_ms, s.active_watched_ms)), 0) / 60000.0)
      ELSE NULL
    END AS tokens_per_min,
    CASE
      WHEN COALESCE(SUM(COALESCE(sm.max_lookups, s.lookup_count)), 0) > 0
        THEN CAST(COALESCE(SUM(COALESCE(sm.max_hits, s.lookup_hits)), 0) AS REAL)
          / CAST(COALESCE(SUM(COALESCE(sm.max_lookups, s.lookup_count)), 0) AS REAL)
      ELSE NULL
    END AS lookup_hit_rate,
    ? AS CREATED_DATE,
    ? AS LAST_UPDATE_DATE
  FROM matching_sessions s
  LEFT JOIN session_metrics sm ON s.session_id = sm.session_id
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
`;

const MONTHLY_ROLLUP_UPSERT = `
  WITH matching_sessions AS (
    SELECT * FROM imm_sessions
    WHERE ${LOCAL_MONTH_EXPR} = ? AND video_id = ?
  ),
  session_metrics AS (
    SELECT
      t.session_id,
      MAX(t.active_watched_ms) AS max_active_ms,
      MAX(t.lines_seen) AS max_lines,
      MAX(t.tokens_seen) AS max_tokens,
      MAX(t.cards_mined) AS max_cards
    FROM imm_session_telemetry t
    JOIN matching_sessions s ON s.session_id = t.session_id
    GROUP BY t.session_id
  )
  INSERT INTO imm_monthly_rollups (
    rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
    total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
  )
  SELECT
    ${LOCAL_MONTH_EXPR.replace('started_at_ms', 's.started_at_ms')} AS rollup_month,
    s.video_id AS video_id,
    COUNT(DISTINCT s.session_id) AS total_sessions,
    COALESCE(SUM(COALESCE(sm.max_active_ms, s.active_watched_ms)), 0) / 60000.0 AS total_active_min,
    COALESCE(SUM(COALESCE(sm.max_lines, s.lines_seen)), 0) AS total_lines_seen,
    COALESCE(SUM(COALESCE(sm.max_tokens, s.tokens_seen)), 0) AS total_tokens_seen,
    COALESCE(SUM(COALESCE(sm.max_cards, s.cards_mined)), 0) AS total_cards,
    ? AS CREATED_DATE,
    ? AS LAST_UPDATE_DATE
  FROM matching_sessions s
  LEFT JOIN session_metrics sm ON s.session_id = sm.session_id
  GROUP BY rollup_month, s.video_id
  ON CONFLICT (rollup_month, video_id) DO UPDATE SET
    total_sessions = excluded.total_sessions,
    total_active_min = excluded.total_active_min,
    total_lines_seen = excluded.total_lines_seen,
    total_tokens_seen = excluded.total_tokens_seen,
    total_cards = excluded.total_cards,
    CREATED_DATE = COALESCE(imm_monthly_rollups.CREATED_DATE, excluded.CREATED_DATE),
    LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
`;

/**
 * Recompute daily/monthly rollup groups touched by the newly merged sessions
 * from the (now merged) local session + telemetry data. The maintenance
 * watermark is left alone: telemetry newer than it gets recomputed again by
 * the app later, which is idempotent.
 */
export function refreshRollupsForNewSessions(
  local: SyncDb,
  newSessionIds: number[],
  summary: SyncMergeSummary,
): void {
  if (newSessionIds.length === 0) return;

  const groups = new Map<string, { day: number; month: number; videoId: number }>();
  for (let offset = 0; offset < newSessionIds.length; offset += 500) {
    const chunk = newSessionIds.slice(offset, offset + 500);
    const rows = selectAll(
      local,
      `SELECT DISTINCT ${LOCAL_DAY_EXPR} AS rollup_day, ${LOCAL_MONTH_EXPR} AS rollup_month, video_id
       FROM imm_sessions WHERE session_id IN (${chunk.map(() => '?').join(',')})`,
      chunk,
    );
    for (const row of rows) {
      const day = Number(row.rollup_day);
      const month = Number(row.rollup_month);
      const videoId = Number(row.video_id);
      groups.set(`${day}-${videoId}`, { day, month, videoId });
    }
  }

  const stampMs = nowDbTimestamp();
  const deleteDaily = local.query(
    'DELETE FROM imm_daily_rollups WHERE rollup_day = ? AND video_id = ?',
  );
  const deleteMonthly = local.query(
    'DELETE FROM imm_monthly_rollups WHERE rollup_month = ? AND video_id = ?',
  );
  const upsertDaily = local.query(DAILY_ROLLUP_UPSERT);
  const upsertMonthly = local.query(MONTHLY_ROLLUP_UPSERT);

  const monthlyGroups = new Set<string>();
  for (const { day, month, videoId } of groups.values()) {
    deleteDaily.run(day, videoId);
    upsertDaily.run(day, videoId, stampMs, stampMs);
    summary.rollupGroupsRecomputed += 1;
    const monthKey = `${month}-${videoId}`;
    if (!monthlyGroups.has(monthKey)) {
      monthlyGroups.add(monthKey);
      deleteMonthly.run(month, videoId);
      upsertMonthly.run(month, videoId, stampMs, stampMs);
    }
  }
}

/**
 * Sessions are pruned after a retention window, but rollups are kept much
 * longer — the remote's older rollup history can't be reconstructed from
 * merged sessions. Copy remote rollup rows for groups where the local DB has
 * neither a rollup row nor any sessions (i.e. history only the remote knows).
 * Groups both machines have data for are never summed, to avoid
 * double-counting sessions that earlier syncs already shared.
 */
export function copyRemoteOnlyRollups(
  local: SyncDb,
  remote: SyncDb,
  videoIdMap: Map<number, number>,
  summary: SyncMergeSummary,
): void {
  if (!tableExists(remote, 'imm_daily_rollups') || !tableExists(local, 'imm_daily_rollups')) return;

  const localDailyExists = local.query(
    'SELECT 1 FROM imm_daily_rollups WHERE rollup_day = ? AND video_id = ? LIMIT 1',
  );
  const localDaySessions = local.query(
    `SELECT 1 FROM imm_sessions WHERE video_id = ? AND ${LOCAL_DAY_EXPR} = ? LIMIT 1`,
  );
  const localMonthSessions = local.query(
    `SELECT 1 FROM imm_sessions WHERE video_id = ? AND ${LOCAL_MONTH_EXPR} = ? LIMIT 1`,
  );
  // rollup_day is a *local* epoch day, so anchor it at local noon (+43200)
  // before reading its month back: plain UTC midnight lands in the previous
  // civil month for the 1st of a month at any negative UTC offset.
  const localMonthSessionsForDay = local.query(
    `SELECT 1 FROM imm_sessions
     WHERE video_id = ?
       AND ${LOCAL_MONTH_EXPR} = CAST(strftime('%Y%m', CAST(? AS INTEGER) * 86400 + 43200, 'unixepoch', 'localtime') AS INTEGER)
     LIMIT 1`,
  );
  const insertDaily = local.query(
    `INSERT INTO imm_daily_rollups (
       rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
       total_tokens_seen, total_cards, cards_per_hour, tokens_per_min, lookup_hit_rate,
       CREATED_DATE, LAST_UPDATE_DATE
     ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
  );
  for (const row of selectAll(remote, 'SELECT * FROM imm_daily_rollups')) {
    if (row.video_id === null) continue;
    const localVideoId = videoIdMap.get(Number(row.video_id));
    if (localVideoId === undefined) continue;
    if (localDailyExists.get(row.rollup_day, localVideoId)) continue;
    if (localDaySessions.get(localVideoId, row.rollup_day)) continue;
    if (localMonthSessionsForDay.get(localVideoId, row.rollup_day)) continue;
    insertDaily.run(
      row.rollup_day,
      localVideoId,
      row.total_sessions,
      row.total_active_min,
      row.total_lines_seen,
      row.total_tokens_seen,
      row.total_cards,
      row.cards_per_hour,
      row.tokens_per_min,
      row.lookup_hit_rate,
      row.CREATED_DATE,
      row.LAST_UPDATE_DATE,
    );
    summary.dailyRollupsCopied += 1;
  }

  const localMonthlyExists = local.query(
    'SELECT 1 FROM imm_monthly_rollups WHERE rollup_month = ? AND video_id = ? LIMIT 1',
  );
  const insertMonthly = local.query(
    `INSERT INTO imm_monthly_rollups (
       rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
       total_tokens_seen, total_cards, CREATED_DATE, LAST_UPDATE_DATE
     ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
  );
  for (const row of selectAll(remote, 'SELECT * FROM imm_monthly_rollups')) {
    if (row.video_id === null) continue;
    const localVideoId = videoIdMap.get(Number(row.video_id));
    if (localVideoId === undefined) continue;
    if (localMonthlyExists.get(row.rollup_month, localVideoId)) continue;
    if (localMonthSessions.get(localVideoId, row.rollup_month)) continue;
    insertMonthly.run(
      row.rollup_month,
      localVideoId,
      row.total_sessions,
      row.total_active_min,
      row.total_lines_seen,
      row.total_tokens_seen,
      row.total_cards,
      row.CREATED_DATE,
      row.LAST_UPDATE_DATE,
    );
    summary.monthlyRollupsCopied += 1;
  }
}
