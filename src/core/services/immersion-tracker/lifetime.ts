import type { DatabaseSync } from './sqlite';
import { finalizeSessionRecord } from './session';
import { nowMs } from './time';
import { forEachIdChunk, makePlaceholders, toDbTimestamp } from './query-shared';
import type { LifetimeRebuildSummary, SessionState } from './types';

interface TelemetryRow {
  active_watched_ms: number | null;
  cards_mined: number | null;
  lines_seen: number | null;
  tokens_seen: number | null;
}

interface VideoRow {
  anime_id: number | null;
  watched: number;
}

interface AnimeRow {
  episodes_total: number | null;
}

function asPositiveNumber(value: number | null, fallback: number): number {
  const resolved = value !== null && Number.isFinite(value) ? value : fallback;
  return Number.isFinite(resolved) ? Math.floor(Math.max(resolved, 0)) : 0;
}

interface ExistenceRow {
  count: number;
}

interface LifetimeMediaStateRow {
  completed: number;
}

interface LifetimeAnimeStateRow {
  episodes_completed: number;
}

interface RetainedSessionRow {
  sessionId: number;
  videoId: number;
  startedAtMs: number | string;
  endedAtMs: number | string;
  lastMediaMs: number | null;
  totalWatchedMs: number;
  activeWatchedMs: number;
  linesSeen: number;
  tokensSeen: number;
  cardsMined: number;
  lookupCount: number;
  lookupHits: number;
  yomitanLookupCount: number;
  pauseCount: number;
  pauseMs: number;
  seekForwardCount: number;
  seekBackwardCount: number;
  mediaBufferEvents: number;
}

const RETAINED_SESSION_METRICS_CTE = `
  retained_sessions AS (
    SELECT
      s.session_id,
      s.video_id,
      v.anime_id,
      s.started_at_ms,
      s.ended_at_ms,
      CAST(MAX(COALESCE(t.active_watched_ms, s.active_watched_ms, 0), 0) AS INTEGER) AS active_ms,
      CAST(MAX(COALESCE(t.cards_mined, s.cards_mined, 0), 0) AS INTEGER) AS cards_mined,
      CAST(MAX(COALESCE(t.lines_seen, s.lines_seen, 0), 0) AS INTEGER) AS lines_seen,
      CAST(MAX(COALESCE(t.tokens_seen, s.tokens_seen, 0), 0) AS INTEGER) AS tokens_seen,
      CASE WHEN v.watched > 0 THEN 1 ELSE 0 END AS completed
    FROM imm_sessions s
    JOIN imm_videos v
      ON v.video_id = s.video_id
    LEFT JOIN imm_session_telemetry t
      ON t.telemetry_id = (
        SELECT telemetry_id
        FROM imm_session_telemetry
        WHERE session_id = s.session_id
        ORDER BY sample_ms DESC, telemetry_id DESC
        LIMIT 1
      )
    WHERE s.ended_at_ms IS NOT NULL
  )
`;

function hasRetainedPriorSession(
  db: DatabaseSync,
  videoId: number,
  startedAtMs: number,
  currentSessionId: number,
): boolean {
  const row = db
    .prepare(
      `
      SELECT 1 AS found
      FROM imm_sessions
      WHERE video_id = ?
        AND (
          CAST(started_at_ms AS REAL) < CAST(? AS REAL)
          OR (
            CAST(started_at_ms AS REAL) = CAST(? AS REAL)
            AND session_id < ?
          )
        )
      LIMIT 1
    `,
    )
    .get(videoId, toDbTimestamp(startedAtMs), toDbTimestamp(startedAtMs), currentSessionId) as {
    found: number;
  } | null;
  return Boolean(row);
}

function isFirstSessionForLocalDay(
  db: DatabaseSync,
  currentSessionId: number,
  startedAtMs: number,
): boolean {
  const row = db
    .prepare(
      `
      SELECT 1 AS found
      FROM imm_sessions
      WHERE session_id != ?
        AND CAST(
          julianday(CAST(started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5
          AS INTEGER
        ) = CAST(
          julianday(CAST(? AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5
          AS INTEGER
        )
        AND (
          CAST(started_at_ms AS REAL) < CAST(? AS REAL)
          OR (
            CAST(started_at_ms AS REAL) = CAST(? AS REAL)
            AND session_id < ?
          )
        )
      LIMIT 1
    `,
    )
    .get(
      currentSessionId,
      toDbTimestamp(startedAtMs),
      toDbTimestamp(startedAtMs),
      toDbTimestamp(startedAtMs),
      currentSessionId,
    ) as { found: number } | null;
  return !row;
}

function resetLifetimeSummaries(db: DatabaseSync, nowMs: number): void {
  db.exec(`
    DELETE FROM imm_lifetime_anime;
    DELETE FROM imm_lifetime_media;
    DELETE FROM imm_lifetime_applied_sessions;
  `);
  db.prepare(
    `
    UPDATE imm_lifetime_global
    SET
      total_sessions = 0,
      total_active_ms = 0,
      total_cards = 0,
      active_days = 0,
      episodes_started = 0,
      episodes_completed = 0,
      anime_completed = 0,
      last_rebuilt_ms = ?,
      LAST_UPDATE_DATE = ?
    WHERE global_id = 1
    `,
  ).run(toDbTimestamp(nowMs), toDbTimestamp(nowMs));
}

function rebuildLifetimeSummariesInternal(
  db: DatabaseSync,
  rebuiltAtMs: number,
): LifetimeRebuildSummary {
  const rebuiltAtDbMs = toDbTimestamp(rebuiltAtMs);
  const appliedSessions = Number(
    (
      db
        .prepare('SELECT COUNT(*) AS total FROM imm_sessions WHERE ended_at_ms IS NOT NULL')
        .get() as { total: number }
    ).total,
  );

  resetLifetimeSummaries(db, rebuiltAtMs);

  db.prepare(
    `
    INSERT INTO imm_lifetime_applied_sessions (
      session_id,
      applied_at_ms,
      CREATED_DATE,
      LAST_UPDATE_DATE
    )
    SELECT
      session_id,
      ended_at_ms,
      ?,
      ?
    FROM imm_sessions
    WHERE ended_at_ms IS NOT NULL
    `,
  ).run(rebuiltAtDbMs, rebuiltAtDbMs);

  db.prepare(
    `
    WITH ${RETAINED_SESSION_METRICS_CTE}
    INSERT INTO imm_lifetime_media (
      video_id,
      total_sessions,
      total_active_ms,
      total_cards,
      total_lines_seen,
      total_tokens_seen,
      completed,
      first_watched_ms,
      last_watched_ms,
      CREATED_DATE,
      LAST_UPDATE_DATE
    )
    SELECT
      video_id,
      COUNT(*) AS total_sessions,
      COALESCE(SUM(active_ms), 0) AS total_active_ms,
      COALESCE(SUM(cards_mined), 0) AS total_cards,
      COALESCE(SUM(lines_seen), 0) AS total_lines_seen,
      COALESCE(SUM(tokens_seen), 0) AS total_tokens_seen,
      MAX(completed) AS completed,
      MIN(started_at_ms) AS first_watched_ms,
      MAX(ended_at_ms) AS last_watched_ms,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM retained_sessions
    GROUP BY video_id
    `,
  ).run(rebuiltAtDbMs, rebuiltAtDbMs);

  db.prepare(
    `
    WITH ${RETAINED_SESSION_METRICS_CTE}
    INSERT INTO imm_lifetime_anime (
      anime_id,
      total_sessions,
      total_active_ms,
      total_cards,
      total_lines_seen,
      total_tokens_seen,
      episodes_started,
      episodes_completed,
      first_watched_ms,
      last_watched_ms,
      CREATED_DATE,
      LAST_UPDATE_DATE
    )
    SELECT
      anime_id,
      COUNT(*) AS total_sessions,
      COALESCE(SUM(active_ms), 0) AS total_active_ms,
      COALESCE(SUM(cards_mined), 0) AS total_cards,
      COALESCE(SUM(lines_seen), 0) AS total_lines_seen,
      COALESCE(SUM(tokens_seen), 0) AS total_tokens_seen,
      COUNT(DISTINCT video_id) AS episodes_started,
      COUNT(DISTINCT CASE WHEN completed > 0 THEN video_id END) AS episodes_completed,
      MIN(started_at_ms) AS first_watched_ms,
      MAX(ended_at_ms) AS last_watched_ms,
      ? AS CREATED_DATE,
      ? AS LAST_UPDATE_DATE
    FROM retained_sessions
    WHERE anime_id IS NOT NULL
    GROUP BY anime_id
    `,
  ).run(rebuiltAtDbMs, rebuiltAtDbMs);

  db.prepare(
    `
    WITH ${RETAINED_SESSION_METRICS_CTE},
    anime_completion AS (
      SELECT
        rs.anime_id,
        MAX(a.episodes_total) AS episodes_total,
        COUNT(DISTINCT CASE WHEN rs.completed > 0 THEN rs.video_id END) AS completed_videos
      FROM retained_sessions rs
      JOIN imm_anime a
        ON a.anime_id = rs.anime_id
      WHERE rs.anime_id IS NOT NULL
      GROUP BY rs.anime_id
    )
    UPDATE imm_lifetime_global
    SET
      total_sessions = (SELECT COUNT(*) FROM retained_sessions),
      total_active_ms = (SELECT COALESCE(SUM(active_ms), 0) FROM retained_sessions),
      total_cards = (SELECT COALESCE(SUM(cards_mined), 0) FROM retained_sessions),
      active_days = (
        SELECT COUNT(DISTINCT CAST(
          julianday(CAST(started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5
          AS INTEGER
        ))
        FROM retained_sessions
      ),
      episodes_started = (SELECT COUNT(DISTINCT video_id) FROM retained_sessions),
      episodes_completed = (
        SELECT COUNT(DISTINCT CASE WHEN completed > 0 THEN video_id END)
        FROM retained_sessions
      ),
      anime_completed = (
        SELECT COUNT(*)
        FROM anime_completion
        WHERE episodes_total IS NOT NULL
          AND episodes_total > 0
          AND completed_videos >= episodes_total
      ),
      last_rebuilt_ms = ?,
      LAST_UPDATE_DATE = ?
    WHERE global_id = 1
    `,
  ).run(rebuiltAtDbMs, rebuiltAtDbMs);

  return {
    appliedSessions,
    rebuiltAtMs,
  };
}

function toRebuildSessionState(row: RetainedSessionRow): SessionState {
  return {
    sessionId: row.sessionId,
    videoId: row.videoId,
    startedAtMs: row.startedAtMs as number,
    currentLineIndex: 0,
    lastWallClockMs: row.endedAtMs as number,
    lastMediaMs: row.lastMediaMs,
    lastPauseStartMs: null,
    isPaused: false,
    pendingTelemetry: false,
    markedWatched: false,
    totalWatchedMs: Math.max(0, row.totalWatchedMs),
    activeWatchedMs: Math.max(0, row.activeWatchedMs),
    linesSeen: Math.max(0, row.linesSeen),
    tokensSeen: Math.max(0, row.tokensSeen),
    cardsMined: Math.max(0, row.cardsMined),
    lookupCount: Math.max(0, row.lookupCount),
    lookupHits: Math.max(0, row.lookupHits),
    yomitanLookupCount: Math.max(0, row.yomitanLookupCount),
    pauseCount: Math.max(0, row.pauseCount),
    pauseMs: Math.max(0, row.pauseMs),
    seekForwardCount: Math.max(0, row.seekForwardCount),
    seekBackwardCount: Math.max(0, row.seekBackwardCount),
    mediaBufferEvents: Math.max(0, row.mediaBufferEvents),
  };
}

function getRetainedStaleActiveSessions(db: DatabaseSync): RetainedSessionRow[] {
  const rows = db
    .prepare(
      `
      SELECT
        s.session_id AS sessionId,
        s.video_id AS videoId,
        s.started_at_ms AS startedAtMs,
        COALESCE(t.sample_ms, s.LAST_UPDATE_DATE, s.started_at_ms) AS endedAtMs,
        s.ended_media_ms AS lastMediaMs,
        COALESCE(t.total_watched_ms, s.total_watched_ms, 0) AS totalWatchedMs,
        COALESCE(t.active_watched_ms, s.active_watched_ms, 0) AS activeWatchedMs,
        COALESCE(t.lines_seen, s.lines_seen, 0) AS linesSeen,
        COALESCE(t.tokens_seen, s.tokens_seen, 0) AS tokensSeen,
        COALESCE(t.cards_mined, s.cards_mined, 0) AS cardsMined,
        COALESCE(t.lookup_count, s.lookup_count, 0) AS lookupCount,
        COALESCE(t.lookup_hits, s.lookup_hits, 0) AS lookupHits,
        COALESCE(t.yomitan_lookup_count, s.yomitan_lookup_count, 0) AS yomitanLookupCount,
        COALESCE(t.pause_count, s.pause_count, 0) AS pauseCount,
        COALESCE(t.pause_ms, s.pause_ms, 0) AS pauseMs,
        COALESCE(t.seek_forward_count, s.seek_forward_count, 0) AS seekForwardCount,
        COALESCE(t.seek_backward_count, s.seek_backward_count, 0) AS seekBackwardCount,
        COALESCE(t.media_buffer_events, s.media_buffer_events, 0) AS mediaBufferEvents
      FROM imm_sessions s
      LEFT JOIN imm_session_telemetry t
        ON t.telemetry_id = (
          SELECT telemetry_id
          FROM imm_session_telemetry
          WHERE session_id = s.session_id
          ORDER BY sample_ms DESC, telemetry_id DESC
          LIMIT 1
        )
      WHERE s.ended_at_ms IS NULL
      ORDER BY s.started_at_ms ASC, s.session_id ASC
      `,
    )
    .all() as Array<
    Omit<RetainedSessionRow, 'startedAtMs' | 'endedAtMs' | 'lastMediaMs'> & {
      startedAtMs: number | string;
      endedAtMs: number | string;
      lastMediaMs: number | string | null;
    }
  >;
  return rows.map((row) => ({
    ...row,
    startedAtMs: row.startedAtMs,
    endedAtMs: row.endedAtMs,
    lastMediaMs: row.lastMediaMs === null ? null : Number(row.lastMediaMs),
  })) as RetainedSessionRow[];
}

function upsertLifetimeMedia(
  db: DatabaseSync,
  videoId: number,
  nowMs: number | string,
  activeMs: number,
  cardsMined: number,
  linesSeen: number,
  tokensSeen: number,
  completed: number,
  startedAtMs: number | string,
  endedAtMs: number | string,
): void {
  db.prepare(
    `
    INSERT INTO imm_lifetime_media(
      video_id,
      total_sessions,
      total_active_ms,
      total_cards,
      total_lines_seen,
      total_tokens_seen,
      completed,
      first_watched_ms,
      last_watched_ms,
      CREATED_DATE,
      LAST_UPDATE_DATE
    )
    VALUES (?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(video_id) DO UPDATE SET
      total_sessions = total_sessions + 1,
      total_active_ms = total_active_ms + excluded.total_active_ms,
      total_cards = total_cards + excluded.total_cards,
      total_lines_seen = total_lines_seen + excluded.total_lines_seen,
      total_tokens_seen = total_tokens_seen + excluded.total_tokens_seen,
      completed = MAX(completed, excluded.completed),
      first_watched_ms = CASE
        WHEN excluded.first_watched_ms IS NULL THEN first_watched_ms
        WHEN first_watched_ms IS NULL THEN excluded.first_watched_ms
        WHEN excluded.first_watched_ms < first_watched_ms THEN excluded.first_watched_ms
        ELSE first_watched_ms
      END,
      last_watched_ms = CASE
        WHEN excluded.last_watched_ms IS NULL THEN last_watched_ms
        WHEN last_watched_ms IS NULL THEN excluded.last_watched_ms
        WHEN excluded.last_watched_ms > last_watched_ms THEN excluded.last_watched_ms
        ELSE last_watched_ms
      END,
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
    `,
  ).run(
    videoId,
    activeMs,
    cardsMined,
    linesSeen,
    tokensSeen,
    completed,
    startedAtMs,
    endedAtMs,
    nowMs,
    nowMs,
  );
}

function upsertLifetimeAnime(
  db: DatabaseSync,
  animeId: number,
  nowMs: number | string,
  activeMs: number,
  cardsMined: number,
  linesSeen: number,
  tokensSeen: number,
  episodesStartedDelta: number,
  episodesCompletedDelta: number,
  startedAtMs: number | string,
  endedAtMs: number | string,
): void {
  db.prepare(
    `
    INSERT INTO imm_lifetime_anime(
      anime_id,
      total_sessions,
      total_active_ms,
      total_cards,
      total_lines_seen,
      total_tokens_seen,
      episodes_started,
      episodes_completed,
      first_watched_ms,
      last_watched_ms,
      CREATED_DATE,
      LAST_UPDATE_DATE
    )
    VALUES (?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(anime_id) DO UPDATE SET
      total_sessions = total_sessions + 1,
      total_active_ms = total_active_ms + excluded.total_active_ms,
      total_cards = total_cards + excluded.total_cards,
      total_lines_seen = total_lines_seen + excluded.total_lines_seen,
      total_tokens_seen = total_tokens_seen + excluded.total_tokens_seen,
      episodes_started = episodes_started + excluded.episodes_started,
      episodes_completed = episodes_completed + excluded.episodes_completed,
      first_watched_ms = CASE
        WHEN excluded.first_watched_ms IS NULL THEN first_watched_ms
        WHEN first_watched_ms IS NULL THEN excluded.first_watched_ms
        WHEN excluded.first_watched_ms < first_watched_ms THEN excluded.first_watched_ms
        ELSE first_watched_ms
      END,
      last_watched_ms = CASE
        WHEN excluded.last_watched_ms IS NULL THEN last_watched_ms
        WHEN last_watched_ms IS NULL THEN excluded.last_watched_ms
        WHEN excluded.last_watched_ms > last_watched_ms THEN excluded.last_watched_ms
        ELSE last_watched_ms
      END,
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
    `,
  ).run(
    animeId,
    activeMs,
    cardsMined,
    linesSeen,
    tokensSeen,
    episodesStartedDelta,
    episodesCompletedDelta,
    startedAtMs,
    endedAtMs,
    nowMs,
    nowMs,
  );
}

export function applySessionLifetimeSummary(
  db: DatabaseSync,
  session: SessionState,
  endedAtMs: number | string,
): void {
  const updatedAtMs = toDbTimestamp(nowMs());
  const applyResult = db
    .prepare(
      `
      INSERT INTO imm_lifetime_applied_sessions (
        session_id,
        applied_at_ms,
        CREATED_DATE,
        LAST_UPDATE_DATE
      ) VALUES (
        ?, ?, ?, ?
      )
      ON CONFLICT(session_id) DO NOTHING
      `,
    )
    .run(session.sessionId, endedAtMs, updatedAtMs, updatedAtMs);

  if ((applyResult.changes ?? 0) <= 0) {
    return;
  }

  const telemetry = db
    .prepare(
      `
    SELECT
      active_watched_ms,
      cards_mined,
      lines_seen,
      tokens_seen
    FROM imm_session_telemetry
    WHERE session_id = ?
    ORDER BY sample_ms DESC, telemetry_id DESC
    LIMIT 1
    `,
    )
    .get(session.sessionId) as TelemetryRow | null;

  const video = db
    .prepare('SELECT anime_id, watched FROM imm_videos WHERE video_id = ?')
    .get(session.videoId) as VideoRow | null;
  const mediaLifetime =
    (db
      .prepare('SELECT completed FROM imm_lifetime_media WHERE video_id = ?')
      .get(session.videoId) as LifetimeMediaStateRow | null | undefined) ?? null;
  const animeLifetime = video?.anime_id
    ? ((db
        .prepare('SELECT episodes_completed FROM imm_lifetime_anime WHERE anime_id = ?')
        .get(video.anime_id) as LifetimeAnimeStateRow | null | undefined) ?? null)
    : null;
  const anime = video?.anime_id
    ? ((db
        .prepare('SELECT episodes_total FROM imm_anime WHERE anime_id = ?')
        .get(video.anime_id) as AnimeRow | null | undefined) ?? null)
    : null;

  const activeMs = asPositiveNumber(telemetry?.active_watched_ms ?? null, session.activeWatchedMs);
  const cardsMined = asPositiveNumber(telemetry?.cards_mined ?? null, session.cardsMined);
  const linesSeen = asPositiveNumber(telemetry?.lines_seen ?? null, session.linesSeen);
  const tokensSeen = asPositiveNumber(telemetry?.tokens_seen ?? null, session.tokensSeen);
  const watched = video?.watched ?? 0;
  const isFirstSessionForVideoRun =
    mediaLifetime === null &&
    !hasRetainedPriorSession(db, session.videoId, session.startedAtMs, session.sessionId);
  const isFirstCompletedSessionForVideoRun =
    watched > 0 && Number(mediaLifetime?.completed ?? 0) <= 0;
  const isFirstSessionForDay = isFirstSessionForLocalDay(
    db,
    session.sessionId,
    session.startedAtMs,
  );
  const episodesCompletedBefore = Number(animeLifetime?.episodes_completed ?? 0);
  const animeEpisodesTotal = anime?.episodes_total ?? null;
  const animeCompletedDelta =
    watched > 0 &&
    isFirstCompletedSessionForVideoRun &&
    animeEpisodesTotal !== null &&
    animeEpisodesTotal > 0 &&
    episodesCompletedBefore < animeEpisodesTotal &&
    episodesCompletedBefore + 1 >= animeEpisodesTotal
      ? 1
      : 0;

  db.prepare(
    `
    UPDATE imm_lifetime_global
    SET
      total_sessions = total_sessions + 1,
      total_active_ms = total_active_ms + ?,
      total_cards = total_cards + ?,
      active_days = active_days + ?,
      episodes_started = episodes_started + ?,
      episodes_completed = episodes_completed + ?,
      anime_completed = anime_completed + ?,
      LAST_UPDATE_DATE = ?
    WHERE global_id = 1
    `,
  ).run(
    activeMs,
    cardsMined,
    isFirstSessionForDay ? 1 : 0,
    isFirstSessionForVideoRun ? 1 : 0,
    isFirstCompletedSessionForVideoRun ? 1 : 0,
    animeCompletedDelta,
    updatedAtMs,
  );

  upsertLifetimeMedia(
    db,
    session.videoId,
    updatedAtMs,
    activeMs,
    cardsMined,
    linesSeen,
    tokensSeen,
    watched > 0 ? 1 : 0,
    session.startedAtMs,
    endedAtMs,
  );

  if (video?.anime_id) {
    upsertLifetimeAnime(
      db,
      video.anime_id,
      updatedAtMs,
      activeMs,
      cardsMined,
      linesSeen,
      tokensSeen,
      isFirstSessionForVideoRun ? 1 : 0,
      isFirstCompletedSessionForVideoRun ? 1 : 0,
      session.startedAtMs,
      endedAtMs,
    );
  }
}

export function rebuildLifetimeSummaries(db: DatabaseSync): LifetimeRebuildSummary {
  const rebuiltAtMs = nowMs();
  db.exec('BEGIN');
  try {
    const summary = rebuildLifetimeSummariesInTransaction(db, rebuiltAtMs);
    db.exec('COMMIT');
    return summary;
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

export function rebuildLifetimeSummariesInTransaction(
  db: DatabaseSync,
  rebuiltAtMs = nowMs(),
): LifetimeRebuildSummary {
  return rebuildLifetimeSummariesInternal(db, rebuiltAtMs);
}

const LOCAL_DAY_EXPR = `CAST(
  julianday(CAST(started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5
  AS INTEGER
)`;

interface LifetimeMediaRemoval {
  videoId: number;
  sessions: number;
  activeMs: number;
  cards: number;
  linesSeen: number;
  tokensSeen: number;
}

/**
 * What a pending delete removes from the lifetime summary tables.
 *
 * Lifetime totals intentionally outlive raw-session retention, so they can
 * never be rebuilt from `imm_sessions` without collapsing history to the
 * retention window. Deletes instead subtract exactly what the deleted rows
 * contributed: this plan is measured before the rows are removed and applied
 * after.
 */
export interface LifetimeRemovalPlan {
  /** Per surviving video: summed metrics of its deleted, lifetime-applied sessions. */
  mediaRemovals: LifetimeMediaRemoval[];
  /** Surviving anime whose lifetime rows must be recomputed from their media rows. */
  affectedAnimeIds: number[];
  /** Local-day keys touched by deleted applied sessions, for active_days upkeep. */
  affectedDayKeys: number[];
}

export function planLifetimeRemovals(
  db: DatabaseSync,
  args: {
    /** Every session being deleted, including ones expanded from video/anime deletes. */
    deletedSessionIds: number[];
    /** Deleted sessions whose video survives the delete. */
    sessionIdsOnSurvivingVideos: number[];
    deletedVideoIds: number[];
    deletedAnimeIds: number[];
  },
): LifetimeRemovalPlan {
  const mediaRemovalsByVideo = new Map<number, LifetimeMediaRemoval>();
  forEachIdChunk(args.sessionIdsOnSurvivingVideos, (chunk) => {
    const rows = db
      .prepare(
        `
        SELECT
          s.video_id AS videoId,
          COUNT(*) AS sessions,
          COALESCE(SUM(CAST(MAX(COALESCE(t.active_watched_ms, s.active_watched_ms, 0), 0) AS INTEGER)), 0) AS activeMs,
          COALESCE(SUM(CAST(MAX(COALESCE(t.cards_mined, s.cards_mined, 0), 0) AS INTEGER)), 0) AS cards,
          COALESCE(SUM(CAST(MAX(COALESCE(t.lines_seen, s.lines_seen, 0), 0) AS INTEGER)), 0) AS linesSeen,
          COALESCE(SUM(CAST(MAX(COALESCE(t.tokens_seen, s.tokens_seen, 0), 0) AS INTEGER)), 0) AS tokensSeen
        FROM imm_sessions s
        JOIN imm_lifetime_applied_sessions a ON a.session_id = s.session_id
        LEFT JOIN imm_session_telemetry t
          ON t.telemetry_id = (
            SELECT telemetry_id
            FROM imm_session_telemetry
            WHERE session_id = s.session_id
            ORDER BY sample_ms DESC, telemetry_id DESC
            LIMIT 1
          )
        WHERE s.session_id IN (${makePlaceholders(chunk)})
        GROUP BY s.video_id
        `,
      )
      .all(...chunk) as LifetimeMediaRemoval[];
    for (const row of rows) {
      const existing = mediaRemovalsByVideo.get(row.videoId);
      if (!existing) {
        mediaRemovalsByVideo.set(row.videoId, { ...row });
        continue;
      }
      existing.sessions += row.sessions;
      existing.activeMs += row.activeMs;
      existing.cards += row.cards;
      existing.linesSeen += row.linesSeen;
      existing.tokensSeen += row.tokensSeen;
    }
  });

  const deletedAnimeIds = new Set(args.deletedAnimeIds);
  const affectedAnimeIds = new Set<number>();
  forEachIdChunk(args.deletedVideoIds, (chunk) => {
    const rows = db
      .prepare(
        `SELECT DISTINCT anime_id AS animeId FROM imm_videos
         WHERE video_id IN (${makePlaceholders(chunk)}) AND anime_id IS NOT NULL`,
      )
      .all(...chunk) as Array<{ animeId: number }>;
    for (const row of rows) affectedAnimeIds.add(row.animeId);
  });
  forEachIdChunk(args.sessionIdsOnSurvivingVideos, (chunk) => {
    const rows = db
      .prepare(
        `SELECT DISTINCT v.anime_id AS animeId
         FROM imm_sessions s
         JOIN imm_videos v ON v.video_id = s.video_id
         WHERE s.session_id IN (${makePlaceholders(chunk)}) AND v.anime_id IS NOT NULL`,
      )
      .all(...chunk) as Array<{ animeId: number }>;
    for (const row of rows) affectedAnimeIds.add(row.animeId);
  });
  for (const animeId of deletedAnimeIds) affectedAnimeIds.delete(animeId);

  const affectedDayKeys = new Set<number>();
  forEachIdChunk(args.deletedSessionIds, (chunk) => {
    const rows = db
      .prepare(
        `SELECT DISTINCT ${LOCAL_DAY_EXPR} AS dayKey
         FROM imm_sessions s
         JOIN imm_lifetime_applied_sessions a ON a.session_id = s.session_id
         WHERE s.session_id IN (${makePlaceholders(chunk)})`,
      )
      .all(...chunk) as Array<{ dayKey: number }>;
    for (const row of rows) affectedDayKeys.add(row.dayKey);
  });

  return {
    mediaRemovals: [...mediaRemovalsByVideo.values()],
    affectedAnimeIds: [...affectedAnimeIds],
    affectedDayKeys: [...affectedDayKeys],
  };
}

/**
 * Apply a removal plan after the underlying rows are gone.
 *
 * Media rows are adjusted by subtraction (pruned-session history stays intact),
 * affected anime rows are recomputed from their surviving media rows, and the
 * global row is re-derived from the media/anime tables. `active_days` is the
 * one metric that can't be derived, so a touched day is only decremented when
 * no ended session remains on that local day; days whose sessions were pruned
 * by retention keep their count because pruning never subtracts.
 */
export function applyLifetimeRemovals(db: DatabaseSync, plan: LifetimeRemovalPlan): void {
  const updatedAtMs = toDbTimestamp(nowMs());

  const subtractMediaStmt = db.prepare(
    `
    UPDATE imm_lifetime_media SET
      total_sessions = MAX(total_sessions - ?, 0),
      total_active_ms = MAX(total_active_ms - ?, 0),
      total_cards = MAX(total_cards - ?, 0),
      total_lines_seen = MAX(total_lines_seen - ?, 0),
      total_tokens_seen = MAX(total_tokens_seen - ?, 0),
      LAST_UPDATE_DATE = ?
    WHERE video_id = ?
    `,
  );
  const dropEmptyMediaStmt = db.prepare(
    'DELETE FROM imm_lifetime_media WHERE video_id = ? AND total_sessions <= 0',
  );
  const remainingSessionRangeStmt = db.prepare(
    `
    SELECT
      MIN(CAST(started_at_ms AS REAL)) AS minStartedMs,
      MAX(CAST(ended_at_ms AS REAL)) AS maxEndedMs
    FROM imm_sessions
    WHERE video_id = ? AND ended_at_ms IS NOT NULL
    `,
  );
  const storedMediaRangeStmt = db.prepare(
    `
    SELECT CAST(first_watched_ms AS REAL) AS firstWatchedMs
    FROM imm_lifetime_media
    WHERE video_id = ?
    `,
  );
  const refreshMediaRangeStmt = db.prepare(
    `
    UPDATE imm_lifetime_media SET
      first_watched_ms = ?,
      last_watched_ms = ?
    WHERE video_id = ?
    `,
  );

  for (const removal of plan.mediaRemovals) {
    subtractMediaStmt.run(
      removal.sessions,
      removal.activeMs,
      removal.cards,
      removal.linesSeen,
      removal.tokensSeen,
      updatedAtMs,
      removal.videoId,
    );
    dropEmptyMediaStmt.run(removal.videoId);
    const stored = storedMediaRangeStmt.get(removal.videoId) as {
      firstWatchedMs: number | null;
    } | null;
    if (!stored) continue;
    const range = remainingSessionRangeStmt.get(removal.videoId) as {
      minStartedMs: number | null;
      maxEndedMs: number | null;
    } | null;
    // Retained sessions are always newer than pruned ones, so the surviving
    // range is authoritative for last_watched while first_watched can only
    // keep or extend the stored (possibly pruned-history) minimum. When no
    // session survives, the stored values are all that's left.
    if (range && range.minStartedMs !== null && range.maxEndedMs !== null) {
      const firstWatchedMs =
        stored.firstWatchedMs === null
          ? range.minStartedMs
          : Math.min(stored.firstWatchedMs, range.minStartedMs);
      refreshMediaRangeStmt.run(
        toDbTimestamp(firstWatchedMs),
        toDbTimestamp(range.maxEndedMs),
        removal.videoId,
      );
    }
  }

  recomputeLifetimeAnimeFromMedia(db, plan.affectedAnimeIds, updatedAtMs);

  // One pass over the sessions rather than a probe per affected day: the local
  // day is a computed expression with no index, so each probe would be a table
  // scan — and the miss case (the day we need to count) is the full-scan one.
  let removedDays = 0;
  if (plan.affectedDayKeys.length > 0) {
    const survivingDayKeys = new Set(
      (
        db
          .prepare(
            `SELECT DISTINCT ${LOCAL_DAY_EXPR} AS dayKey
             FROM imm_sessions
             WHERE ended_at_ms IS NOT NULL`,
          )
          .all() as Array<{ dayKey: number }>
      ).map((row) => row.dayKey),
    );
    for (const dayKey of plan.affectedDayKeys) {
      if (!survivingDayKeys.has(dayKey)) removedDays += 1;
    }
  }

  recomputeLifetimeGlobalFromSummaries(db, { removedActiveDays: removedDays, updatedAtMs });
}

/**
 * Recompute lifetime anime rows exactly from their surviving media rows.
 *
 * Media rows are the durable per-video ledger (they outlive session pruning and
 * follow a video when it moves between anime), so this is the correct refresh
 * after merges, moves, and deletes. Anime with no media rows left are dropped.
 */
export function recomputeLifetimeAnimeFromMedia(
  db: DatabaseSync,
  animeIds: number[],
  updatedAtMs = toDbTimestamp(nowMs()),
): void {
  if (animeIds.length === 0) return;

  const animeSummaryStmt = db.prepare(
    `
    SELECT
      COUNT(*) AS episodeRows,
      COALESCE(SUM(m.total_sessions), 0) AS totalSessions,
      COALESCE(SUM(m.total_active_ms), 0) AS totalActiveMs,
      COALESCE(SUM(m.total_cards), 0) AS totalCards,
      COALESCE(SUM(m.total_lines_seen), 0) AS totalLinesSeen,
      COALESCE(SUM(m.total_tokens_seen), 0) AS totalTokensSeen,
      COALESCE(SUM(CASE WHEN m.completed > 0 THEN 1 ELSE 0 END), 0) AS episodesCompleted,
      MIN(CAST(m.first_watched_ms AS REAL)) AS firstWatchedMs,
      MAX(CAST(m.last_watched_ms AS REAL)) AS lastWatchedMs
    FROM imm_lifetime_media m
    JOIN imm_videos v ON v.video_id = m.video_id
    WHERE v.anime_id = ?
    `,
  );
  const dropAnimeStmt = db.prepare('DELETE FROM imm_lifetime_anime WHERE anime_id = ?');
  const upsertAnimeStmt = db.prepare(
    `
    INSERT INTO imm_lifetime_anime(
      anime_id,
      total_sessions,
      total_active_ms,
      total_cards,
      total_lines_seen,
      total_tokens_seen,
      episodes_started,
      episodes_completed,
      first_watched_ms,
      last_watched_ms,
      CREATED_DATE,
      LAST_UPDATE_DATE
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(anime_id) DO UPDATE SET
      total_sessions = excluded.total_sessions,
      total_active_ms = excluded.total_active_ms,
      total_cards = excluded.total_cards,
      total_lines_seen = excluded.total_lines_seen,
      total_tokens_seen = excluded.total_tokens_seen,
      episodes_started = excluded.episodes_started,
      episodes_completed = excluded.episodes_completed,
      first_watched_ms = excluded.first_watched_ms,
      last_watched_ms = excluded.last_watched_ms,
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
    `,
  );

  for (const animeId of animeIds) {
    const summary = animeSummaryStmt.get(animeId) as {
      episodeRows: number;
      totalSessions: number;
      totalActiveMs: number;
      totalCards: number;
      totalLinesSeen: number;
      totalTokensSeen: number;
      episodesCompleted: number;
      firstWatchedMs: number | null;
      lastWatchedMs: number | null;
    };
    if (Number(summary.episodeRows) === 0) {
      dropAnimeStmt.run(animeId);
      continue;
    }
    upsertAnimeStmt.run(
      animeId,
      summary.totalSessions,
      summary.totalActiveMs,
      summary.totalCards,
      summary.totalLinesSeen,
      summary.totalTokensSeen,
      summary.episodeRows,
      summary.episodesCompleted,
      summary.firstWatchedMs === null ? null : toDbTimestamp(summary.firstWatchedMs),
      summary.lastWatchedMs === null ? null : toDbTimestamp(summary.lastWatchedMs),
      updatedAtMs,
      updatedAtMs,
    );
  }
}

/**
 * Re-derive the global lifetime row from the media/anime summary tables.
 *
 * Every global metric except active_days is a pure aggregate of those tables;
 * active_days can't be derived, so callers pass how many day slots their change
 * removed (0 for moves/merges, which never touch sessions).
 */
export function recomputeLifetimeGlobalFromSummaries(
  db: DatabaseSync,
  options: { removedActiveDays?: number; updatedAtMs?: string } = {},
): void {
  const updatedAtMs = options.updatedAtMs ?? toDbTimestamp(nowMs());
  const mediaTotals = db
    .prepare(
      `
      SELECT
        COUNT(*) AS episodesStarted,
        COALESCE(SUM(total_sessions), 0) AS totalSessions,
        COALESCE(SUM(total_active_ms), 0) AS totalActiveMs,
        COALESCE(SUM(total_cards), 0) AS totalCards,
        COALESCE(SUM(CASE WHEN completed > 0 THEN 1 ELSE 0 END), 0) AS episodesCompleted
      FROM imm_lifetime_media
      `,
    )
    .get() as {
    episodesStarted: number;
    totalSessions: number;
    totalActiveMs: number;
    totalCards: number;
    episodesCompleted: number;
  };
  const animeCompletedRow = db
    .prepare(
      `
      SELECT COUNT(*) AS animeCompleted
      FROM imm_lifetime_anime la
      JOIN imm_anime a ON a.anime_id = la.anime_id
      WHERE a.episodes_total IS NOT NULL
        AND a.episodes_total > 0
        AND la.episodes_completed >= a.episodes_total
      `,
    )
    .get() as { animeCompleted: number };

  db.prepare(
    `
    UPDATE imm_lifetime_global SET
      total_sessions = ?,
      total_active_ms = ?,
      total_cards = ?,
      episodes_started = ?,
      episodes_completed = ?,
      anime_completed = ?,
      active_days = MAX(active_days - ?, 0),
      LAST_UPDATE_DATE = ?
    WHERE global_id = 1
    `,
  ).run(
    mediaTotals.totalSessions,
    mediaTotals.totalActiveMs,
    mediaTotals.totalCards,
    mediaTotals.episodesStarted,
    mediaTotals.episodesCompleted,
    animeCompletedRow.animeCompleted,
    options.removedActiveDays ?? 0,
    updatedAtMs,
  );
}

export interface LifetimeRepairSummary {
  recomputedAnime: number;
  repairedAtMs: number;
}

/**
 * Non-destructive lifetime repair: recompute every anime row and the global row
 * from the per-video media ledger.
 *
 * Unlike {@link rebuildLifetimeSummaries}, this never resets the tables from
 * retained sessions, so lifetime history older than the session retention
 * window survives. The one exception is a database whose lifetime tables were
 * never populated — there is no ledger to repair from, so it bootstraps with
 * the full rebuild instead.
 */
export function repairLifetimeSummariesFromMedia(db: DatabaseSync): LifetimeRepairSummary {
  const repairedAtMs = nowMs();
  let transactionStarted = false;
  try {
    db.exec('BEGIN IMMEDIATE');
    transactionStarted = true;
    if (shouldBackfillLifetimeSummaries(db)) {
      const rebuilt = rebuildLifetimeSummariesInTransaction(db, repairedAtMs);
      const animeRow = db
        .prepare('SELECT COUNT(*) AS count FROM imm_lifetime_anime')
        .get() as ExistenceRow;
      db.exec('COMMIT');
      return { recomputedAnime: Number(animeRow.count), repairedAtMs: rebuilt.rebuiltAtMs };
    }

    const animeIds = new Set<number>();
    for (const row of db
      .prepare('SELECT DISTINCT anime_id AS animeId FROM imm_videos WHERE anime_id IS NOT NULL')
      .all() as Array<{ animeId: number }>) {
      animeIds.add(row.animeId);
    }
    for (const row of db
      .prepare('SELECT anime_id AS animeId FROM imm_lifetime_anime')
      .all() as Array<{ animeId: number }>) {
      animeIds.add(row.animeId);
    }
    const updatedAtMs = toDbTimestamp(repairedAtMs);
    recomputeLifetimeAnimeFromMedia(db, [...animeIds], updatedAtMs);
    recomputeLifetimeGlobalFromSummaries(db, { updatedAtMs });
    db.exec('COMMIT');
    return { recomputedAnime: animeIds.size, repairedAtMs };
  } catch (error) {
    if (transactionStarted) db.exec('ROLLBACK');
    throw error;
  }
}

export function reconcileStaleActiveSessions(db: DatabaseSync): number {
  const sessions = getRetainedStaleActiveSessions(db);
  if (sessions.length === 0) {
    return 0;
  }

  db.exec('BEGIN');
  try {
    for (const session of sessions) {
      const state = toRebuildSessionState(session);
      finalizeSessionRecord(db, state, session.endedAtMs);
      applySessionLifetimeSummary(db, state, session.endedAtMs);
    }
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }

  return sessions.length;
}

export function shouldBackfillLifetimeSummaries(db: DatabaseSync): boolean {
  const globalRow = db
    .prepare('SELECT total_sessions AS totalSessions FROM imm_lifetime_global WHERE global_id = 1')
    .get() as { totalSessions: number } | null;
  const appliedRow = db
    .prepare('SELECT COUNT(*) AS count FROM imm_lifetime_applied_sessions')
    .get() as ExistenceRow | null;
  const endedRow = db
    .prepare('SELECT COUNT(*) AS count FROM imm_sessions WHERE ended_at_ms IS NOT NULL')
    .get() as ExistenceRow | null;

  const totalSessions = Number(globalRow?.totalSessions ?? 0);
  const appliedSessions = Number(appliedRow?.count ?? 0);
  const retainedEndedSessions = Number(endedRow?.count ?? 0);

  return retainedEndedSessions > 0 && (appliedSessions === 0 || totalSessions === 0);
}
