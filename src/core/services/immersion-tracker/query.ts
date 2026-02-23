import type { DatabaseSync } from 'node:sqlite';
import type {
  ImmersionSessionRollupRow,
  SessionSummaryQueryRow,
  SessionTimelineRow,
} from './types';

export function getSessionSummaries(db: DatabaseSync, limit = 50): SessionSummaryQueryRow[] {
  const prepared = db.prepare(`
    SELECT
      s.video_id AS videoId,
      s.started_at_ms AS startedAtMs,
      s.ended_at_ms AS endedAtMs,
      COALESCE(SUM(t.total_watched_ms), 0) AS totalWatchedMs,
      COALESCE(SUM(t.active_watched_ms), 0) AS activeWatchedMs,
      COALESCE(SUM(t.lines_seen), 0) AS linesSeen,
      COALESCE(SUM(t.words_seen), 0) AS wordsSeen,
      COALESCE(SUM(t.tokens_seen), 0) AS tokensSeen,
      COALESCE(SUM(t.cards_mined), 0) AS cardsMined,
      COALESCE(SUM(t.lookup_count), 0) AS lookupCount,
      COALESCE(SUM(t.lookup_hits), 0) AS lookupHits
    FROM imm_sessions s
    LEFT JOIN imm_session_telemetry t ON t.session_id = s.session_id
    GROUP BY s.session_id
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `);
  return prepared.all(limit) as unknown as SessionSummaryQueryRow[];
}

export function getSessionTimeline(
  db: DatabaseSync,
  sessionId: number,
  limit = 200,
): SessionTimelineRow[] {
  const prepared = db.prepare(`
    SELECT
      sample_ms AS sampleMs,
      total_watched_ms AS totalWatchedMs,
      active_watched_ms AS activeWatchedMs,
      lines_seen AS linesSeen,
      words_seen AS wordsSeen,
      tokens_seen AS tokensSeen,
      cards_mined AS cardsMined
    FROM imm_session_telemetry
    WHERE session_id = ?
    ORDER BY sample_ms DESC
    LIMIT ?
  `);
  return prepared.all(sessionId, limit) as unknown as SessionTimelineRow[];
}

export function getQueryHints(db: DatabaseSync): {
  totalSessions: number;
  activeSessions: number;
} {
  const sessions = db.prepare('SELECT COUNT(*) AS total FROM imm_sessions');
  const active = db.prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE ended_at_ms IS NULL');
  const totalSessions = Number(sessions.get()?.total ?? 0);
  const activeSessions = Number(active.get()?.total ?? 0);
  return { totalSessions, activeSessions };
}

export function getDailyRollups(db: DatabaseSync, limit = 60): ImmersionSessionRollupRow[] {
  const prepared = db.prepare(`
    SELECT
      rollup_day AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      total_words_seen AS totalWordsSeen,
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      cards_per_hour AS cardsPerHour,
      words_per_min AS wordsPerMin,
      lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups
    ORDER BY rollup_day DESC, video_id DESC
    LIMIT ?
  `);
  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}

export function getMonthlyRollups(db: DatabaseSync, limit = 24): ImmersionSessionRollupRow[] {
  const prepared = db.prepare(`
    SELECT
      rollup_month AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      total_words_seen AS totalWordsSeen,
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      0 AS cardsPerHour,
      0 AS wordsPerMin,
      0 AS lookupHitRate
    FROM imm_monthly_rollups
    ORDER BY rollup_month DESC, video_id DESC
    LIMIT ?
  `);
  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}
