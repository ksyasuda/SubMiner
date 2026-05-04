import type { DatabaseSync } from './sqlite';
import type {
  ImmersionSessionRollupRow,
  SessionSummaryQueryRow,
  SessionTimelineRow,
} from './types';
import {
  ACTIVE_SESSION_METRICS_CTE,
  SESSION_WORD_COUNTS_CTE,
  SESSION_WORD_COUNTS_SELECT,
  currentDbTimestamp,
  fromDbTimestamp,
  getLocalEpochDay,
  getShiftedLocalDaySec,
  sessionDisplayWordsExpr,
  visibleWordSql,
} from './query-shared';

export function getSessionSummaries(db: DatabaseSync, limit = 50): SessionSummaryQueryRow[] {
  const prepared = db.prepare(`
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      s.session_id AS sessionId,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      v.anime_id AS animeId,
      a.canonical_title AS animeTitle,
      s.started_at_ms AS startedAtMs,
      s.ended_at_ms AS endedAtMs,
      COALESCE(asm.totalWatchedMs, s.total_watched_ms, 0) AS totalWatchedMs,
      COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0) AS activeWatchedMs,
      COALESCE(asm.linesSeen, s.lines_seen, 0) AS linesSeen,
      COALESCE(asm.tokensSeen, s.tokens_seen, 0) AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.lookupCount, s.lookup_count, 0) AS lookupCount,
      COALESCE(asm.lookupHits, s.lookup_hits, 0) AS lookupHits,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount
    FROM imm_sessions s
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `);
  const rows = prepared.all(limit) as Array<
    SessionSummaryQueryRow & {
      startedAtMs: number | string;
      endedAtMs: number | string | null;
    }
  >;
  return rows.map((row) => ({
    ...row,
    startedAtMs: fromDbTimestamp(row.startedAtMs) ?? 0,
    endedAtMs: fromDbTimestamp(row.endedAtMs),
  }));
}

export function getSessionTimeline(
  db: DatabaseSync,
  sessionId: number,
  limit?: number,
): SessionTimelineRow[] {
  const select = `
    SELECT
      sample_ms AS sampleMs,
      total_watched_ms AS totalWatchedMs,
      active_watched_ms AS activeWatchedMs,
      lines_seen AS linesSeen,
      tokens_seen AS tokensSeen,
      cards_mined AS cardsMined
    FROM imm_session_telemetry
    WHERE session_id = ?
    ORDER BY sample_ms DESC, telemetry_id DESC
  `;

  if (limit === undefined) {
    const rows = db.prepare(select).all(sessionId) as Array<
      SessionTimelineRow & {
        sampleMs: number | string;
      }
    >;
    return rows.map((row) => ({
      ...row,
      sampleMs: fromDbTimestamp(row.sampleMs) ?? 0,
    }));
  }
  const rows = db.prepare(`${select}\n    LIMIT ?`).all(sessionId, limit) as Array<
    SessionTimelineRow & {
      sampleMs: number | string;
    }
  >;
  return rows.map((row) => ({
    ...row,
    sampleMs: fromDbTimestamp(row.sampleMs) ?? 0,
  }));
}

/** Returns all distinct headwords in the vocabulary table (global). */
export function getAllDistinctHeadwords(db: DatabaseSync): string[] {
  const rows = db
    .prepare(`SELECT DISTINCT headword FROM imm_words w WHERE ${visibleWordSql('w')}`)
    .all() as Array<{
    headword: string;
  }>;
  return rows.map((r) => r.headword);
}

/** Returns distinct headwords seen for a specific anime. */
export function getAnimeDistinctHeadwords(db: DatabaseSync, animeId: number): string[] {
  const rows = db
    .prepare(
      `
      SELECT DISTINCT w.headword
      FROM imm_word_line_occurrences o
      JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
      JOIN imm_words w ON w.id = o.word_id
      WHERE sl.anime_id = ? AND ${visibleWordSql('w')}
    `,
    )
    .all(animeId) as Array<{ headword: string }>;
  return rows.map((r) => r.headword);
}

/** Returns distinct headwords seen for a specific video/media. */
export function getMediaDistinctHeadwords(db: DatabaseSync, videoId: number): string[] {
  const rows = db
    .prepare(
      `
      SELECT DISTINCT w.headword
      FROM imm_word_line_occurrences o
      JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
      JOIN imm_words w ON w.id = o.word_id
      WHERE sl.video_id = ? AND ${visibleWordSql('w')}
    `,
    )
    .all(videoId) as Array<{ headword: string }>;
  return rows.map((r) => r.headword);
}

/**
 * Returns the headword for each word seen in a session, grouped by line_index.
 * Used to compute cumulative known-words counts for the session timeline chart.
 */
export function getSessionWordsByLine(
  db: DatabaseSync,
  sessionId: number,
): Array<{ lineIndex: number; headword: string; occurrenceCount: number }> {
  const stmt = db.prepare(`
    SELECT
      sl.line_index AS lineIndex,
      w.headword AS headword,
      wlo.occurrence_count AS occurrenceCount
    FROM imm_subtitle_lines sl
    JOIN imm_word_line_occurrences wlo ON wlo.line_id = sl.line_id
    JOIN imm_words w ON w.id = wlo.word_id
    WHERE sl.session_id = ? AND ${visibleWordSql('w')}
    ORDER BY sl.line_index ASC
  `);
  return stmt.all(sessionId) as Array<{
    lineIndex: number;
    headword: string;
    occurrenceCount: number;
  }>;
}

function getNewWordCounts(db: DatabaseSync): { newWordsToday: number; newWordsThisWeek: number } {
  const currentTimestamp = currentDbTimestamp();
  const todayStartSec = getShiftedLocalDaySec(db, currentTimestamp, 0);
  const weekAgoSec = getShiftedLocalDaySec(db, currentTimestamp, -7);

  const rows = db
    .prepare(
      `
    SELECT
      headword,
      first_seen AS firstSeen
    FROM imm_words
    WHERE first_seen IS NOT NULL
      AND headword IS NOT NULL
      AND headword != ''
  `,
    )
    .all() as Array<{ headword: string; firstSeen: number | string }>;

  const firstSeenByHeadword = new Map<string, number>();
  for (const row of rows) {
    const firstSeen = Number(row.firstSeen);
    if (!Number.isFinite(firstSeen)) {
      continue;
    }
    const previous = firstSeenByHeadword.get(row.headword);
    if (previous === undefined || firstSeen < previous) {
      firstSeenByHeadword.set(row.headword, firstSeen);
    }
  }

  let today = 0;
  let week = 0;
  for (const firstSeen of firstSeenByHeadword.values()) {
    if (firstSeen >= todayStartSec) {
      today += 1;
    }
    if (firstSeen >= weekAgoSec) {
      week += 1;
    }
  }

  return {
    newWordsToday: today,
    newWordsThisWeek: week,
  };
}

export function getQueryHints(db: DatabaseSync): {
  totalSessions: number;
  activeSessions: number;
  episodesToday: number;
  activeAnimeCount: number;
  totalEpisodesWatched: number;
  totalAnimeCompleted: number;
  totalActiveMin: number;
  totalCards: number;
  activeDays: number;
  totalTokensSeen: number;
  totalLookupCount: number;
  totalLookupHits: number;
  totalYomitanLookupCount: number;
  newWordsToday: number;
  newWordsThisWeek: number;
} {
  const active = db.prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE ended_at_ms IS NULL');
  const activeSessions = Number((active.get() as { total?: number } | null)?.total ?? 0);
  const lifetime = db
    .prepare(
      `
    SELECT
      total_sessions AS totalSessions,
      total_active_ms AS totalActiveMs,
      total_cards AS totalCards,
      active_days AS activeDays,
      episodes_completed AS episodesCompleted,
      anime_completed AS animeCompleted
    FROM imm_lifetime_global
    WHERE global_id = 1
  `,
    )
    .get() as {
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    activeDays: number;
    episodesCompleted: number;
    animeCompleted: number;
  } | null;

  const currentTimestamp = currentDbTimestamp();
  const todayLocal = getLocalEpochDay(db, currentTimestamp);

  const episodesToday =
    (
      db
        .prepare(
          `
    SELECT COUNT(DISTINCT s.video_id) AS count
    FROM imm_sessions s
    WHERE CAST(
      julianday(CAST(s.started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5
      AS INTEGER
    ) = ?
  `,
        )
        .get(todayLocal) as { count: number }
    )?.count ?? 0;

  const thirtyDaysAgoMs = getShiftedLocalDaySec(db, currentTimestamp, -30).toString() + '000';
  const activeAnimeCount =
    (
      db
        .prepare(
          `
    SELECT COUNT(DISTINCT v.anime_id) AS count
    FROM imm_sessions s
    JOIN imm_videos v ON v.video_id = s.video_id
    WHERE v.anime_id IS NOT NULL
    AND s.started_at_ms >= ?
  `,
        )
        .get(thirtyDaysAgoMs) as { count: number }
    )?.count ?? 0;

  const totalEpisodesWatched = Number(lifetime?.episodesCompleted ?? 0);
  const totalAnimeCompleted = Number(lifetime?.animeCompleted ?? 0);
  const totalSessions = Number(lifetime?.totalSessions ?? 0);
  const totalActiveMin = Math.floor(Math.max(0, lifetime?.totalActiveMs ?? 0) / 60000);
  const totalCards = Number(lifetime?.totalCards ?? 0);
  const activeDays = Number(lifetime?.activeDays ?? 0);

  const lookupWordsExpr = sessionDisplayWordsExpr(
    's',
    'swc',
    'COALESCE(t.tokens_seen, s.tokens_seen)',
  );
  const lookupTotals = db
    .prepare(
      `
    ${SESSION_WORD_COUNTS_CTE}
    SELECT
      COALESCE(SUM(${lookupWordsExpr}), 0) AS totalTokensSeen,
      COALESCE(SUM(COALESCE(t.lookup_count, s.lookup_count, 0)), 0) AS totalLookupCount,
      COALESCE(SUM(COALESCE(t.lookup_hits, s.lookup_hits, 0)), 0) AS totalLookupHits,
      COALESCE(SUM(COALESCE(t.yomitan_lookup_count, s.yomitan_lookup_count, 0)), 0) AS totalYomitanLookupCount
    FROM imm_sessions s
    LEFT JOIN (
      SELECT
        session_id,
        MAX(tokens_seen) AS tokens_seen,
        MAX(lookup_count) AS lookup_count,
        MAX(lookup_hits) AS lookup_hits,
        MAX(yomitan_lookup_count) AS yomitan_lookup_count
      FROM imm_session_telemetry
      GROUP BY session_id
    ) t ON t.session_id = s.session_id
    LEFT JOIN session_word_counts swc ON swc.sessionId = s.session_id
    WHERE s.ended_at_ms IS NOT NULL
  `,
    )
    .get() as {
    totalTokensSeen: number;
    totalLookupCount: number;
    totalLookupHits: number;
    totalYomitanLookupCount: number;
  } | null;

  return {
    totalSessions,
    activeSessions,
    episodesToday,
    activeAnimeCount,
    totalEpisodesWatched,
    totalAnimeCompleted,
    totalActiveMin,
    totalCards,
    activeDays,
    totalTokensSeen: Number(lookupTotals?.totalTokensSeen ?? 0),
    totalLookupCount: Number(lookupTotals?.totalLookupCount ?? 0),
    totalLookupHits: Number(lookupTotals?.totalLookupHits ?? 0),
    totalYomitanLookupCount: Number(lookupTotals?.totalYomitanLookupCount ?? 0),
    ...getNewWordCounts(db),
  };
}

export function getDailyRollups(db: DatabaseSync, limit = 60): ImmersionSessionRollupRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc');
  const prepared = db.prepare(`
    WITH session_word_counts AS (
      ${SESSION_WORD_COUNTS_SELECT}
    ),
    daily_word_counts AS (
      SELECT
        CAST(
          julianday(CAST(s.started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5
          AS INTEGER
        ) AS rollupDay,
        s.video_id AS videoId,
        SUM(${wordsExpr}) AS totalTokensSeen
      FROM imm_sessions s
      LEFT JOIN session_word_counts swc ON swc.sessionId = s.session_id
      WHERE s.ended_at_ms IS NOT NULL
      GROUP BY rollupDay, s.video_id
    ),
    recent_days AS (
      SELECT DISTINCT rollup_day
      FROM imm_daily_rollups
      ORDER BY rollup_day DESC
      LIMIT ?
    )
    SELECT
      r.rollup_day AS rollupDayOrMonth,
      r.video_id AS videoId,
      r.total_sessions AS totalSessions,
      r.total_active_min AS totalActiveMin,
      r.total_lines_seen AS totalLinesSeen,
      CASE
        WHEN dwc.totalTokensSeen IS NOT NULL AND dwc.totalTokensSeen > r.total_tokens_seen THEN dwc.totalTokensSeen
        ELSE r.total_tokens_seen
      END AS totalTokensSeen,
      r.total_cards AS totalCards,
      r.cards_per_hour AS cardsPerHour,
      CASE
        WHEN r.total_active_min > 0 THEN (
          CASE
            WHEN dwc.totalTokensSeen IS NOT NULL AND dwc.totalTokensSeen > r.total_tokens_seen THEN dwc.totalTokensSeen
            ELSE r.total_tokens_seen
          END
        ) * 1.0 / r.total_active_min
        ELSE NULL
      END AS tokensPerMin,
      r.lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups r
    LEFT JOIN daily_word_counts dwc
      ON dwc.rollupDay = r.rollup_day
      AND (
        (dwc.videoId IS NULL AND r.video_id IS NULL)
        OR dwc.videoId = r.video_id
      )
    WHERE r.rollup_day IN (SELECT rollup_day FROM recent_days)
    ORDER BY r.rollup_day DESC, r.video_id DESC
    `);

  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}

export function getMonthlyRollups(db: DatabaseSync, limit = 24): ImmersionSessionRollupRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc');
  const prepared = db.prepare(`
    WITH session_word_counts AS (
      ${SESSION_WORD_COUNTS_SELECT}
    ),
    monthly_word_counts AS (
      SELECT
        CAST(strftime('%Y%m', CAST(s.started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') AS INTEGER) AS rollupMonth,
        s.video_id AS videoId,
        SUM(${wordsExpr}) AS totalTokensSeen
      FROM imm_sessions s
      LEFT JOIN session_word_counts swc ON swc.sessionId = s.session_id
      WHERE s.ended_at_ms IS NOT NULL
      GROUP BY rollupMonth, s.video_id
    ),
    recent_months AS (
      SELECT DISTINCT rollup_month
      FROM imm_monthly_rollups
      ORDER BY rollup_month DESC
      LIMIT ?
    )
    SELECT
      r.rollup_month AS rollupDayOrMonth,
      r.video_id AS videoId,
      r.total_sessions AS totalSessions,
      r.total_active_min AS totalActiveMin,
      r.total_lines_seen AS totalLinesSeen,
      CASE
        WHEN mwc.totalTokensSeen IS NOT NULL AND mwc.totalTokensSeen > r.total_tokens_seen THEN mwc.totalTokensSeen
        ELSE r.total_tokens_seen
      END AS totalTokensSeen,
      r.total_cards AS totalCards,
      CASE
        WHEN r.total_active_min > 0 THEN (r.total_cards * 60.0) / r.total_active_min
        ELSE NULL
      END AS cardsPerHour,
      CASE
        WHEN r.total_active_min > 0 THEN (
          CASE
            WHEN mwc.totalTokensSeen IS NOT NULL AND mwc.totalTokensSeen > r.total_tokens_seen THEN mwc.totalTokensSeen
            ELSE r.total_tokens_seen
          END
        ) * 1.0 / r.total_active_min
        ELSE NULL
      END AS tokensPerMin,
      NULL AS lookupHitRate
    FROM imm_monthly_rollups r
    LEFT JOIN monthly_word_counts mwc
      ON mwc.rollupMonth = r.rollup_month
      AND (
        (mwc.videoId IS NULL AND r.video_id IS NULL)
        OR mwc.videoId = r.video_id
      )
    WHERE r.rollup_month IN (SELECT rollup_month FROM recent_months)
    ORDER BY r.rollup_month DESC, r.video_id DESC
  `);
  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}
