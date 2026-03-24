import type { DatabaseSync } from './sqlite';
import { finalizeSessionRecord } from './session';
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
  if (value === null || !Number.isFinite(value)) {
    return fallback;
  }
  return Math.max(0, Math.floor(value));
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
  startedAtMs: number;
  endedAtMs: number;
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

function hasRetainedPriorSession(
  db: DatabaseSync,
  videoId: number,
  startedAtMs: number,
  currentSessionId: number,
): boolean {
  return (
    Number(
      (
        db
          .prepare(
            `
            SELECT COUNT(*) AS count
            FROM imm_sessions
            WHERE video_id = ?
              AND (
                started_at_ms < ?
                OR (started_at_ms = ? AND session_id < ?)
              )
            `,
          )
          .get(videoId, startedAtMs, startedAtMs, currentSessionId) as ExistenceRow | null
      )?.count ?? 0,
    ) > 0
  );
}

function isFirstSessionForLocalDay(
  db: DatabaseSync,
  currentSessionId: number,
  startedAtMs: number,
): boolean {
  return (
    (
      db
        .prepare(
          `
      SELECT COUNT(*) AS count
      FROM imm_sessions
      WHERE CAST(strftime('%s', started_at_ms / 1000, 'unixepoch', 'localtime') AS INTEGER) / 86400
        = CAST(strftime('%s', ? / 1000, 'unixepoch', 'localtime') AS INTEGER) / 86400
        AND (
          started_at_ms < ?
          OR (started_at_ms = ? AND session_id < ?)
        )
      `,
        )
        .get(startedAtMs, startedAtMs, startedAtMs, currentSessionId) as ExistenceRow | null
    )?.count === 0
  );
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
  ).run(nowMs, nowMs);
}

function rebuildLifetimeSummariesInternal(
  db: DatabaseSync,
  rebuiltAtMs: number,
): LifetimeRebuildSummary {
  const sessions = db
    .prepare(
      `
      SELECT
        session_id AS sessionId,
        video_id AS videoId,
        started_at_ms AS startedAtMs,
        ended_at_ms AS endedAtMs,
        total_watched_ms AS totalWatchedMs,
        active_watched_ms AS activeWatchedMs,
        lines_seen AS linesSeen,
        tokens_seen AS tokensSeen,
        cards_mined AS cardsMined,
        lookup_count AS lookupCount,
        lookup_hits AS lookupHits,
        yomitan_lookup_count AS yomitanLookupCount,
        pause_count AS pauseCount,
        pause_ms AS pauseMs,
        seek_forward_count AS seekForwardCount,
        seek_backward_count AS seekBackwardCount,
        media_buffer_events AS mediaBufferEvents
      FROM imm_sessions
      WHERE ended_at_ms IS NOT NULL
      ORDER BY started_at_ms ASC, session_id ASC
      `,
    )
    .all() as RetainedSessionRow[];

  resetLifetimeSummaries(db, rebuiltAtMs);
  for (const session of sessions) {
    applySessionLifetimeSummary(db, toRebuildSessionState(session), session.endedAtMs);
  }

  return {
    appliedSessions: sessions.length,
    rebuiltAtMs,
  };
}

function toRebuildSessionState(row: RetainedSessionRow): SessionState {
  return {
    sessionId: row.sessionId,
    videoId: row.videoId,
    startedAtMs: row.startedAtMs,
    currentLineIndex: 0,
    lastWallClockMs: row.endedAtMs,
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
  return db
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
    .all() as RetainedSessionRow[];
}

function upsertLifetimeMedia(
  db: DatabaseSync,
  videoId: number,
  nowMs: number,
  activeMs: number,
  cardsMined: number,
  linesSeen: number,
  tokensSeen: number,
  completed: number,
  startedAtMs: number,
  endedAtMs: number,
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
  nowMs: number,
  activeMs: number,
  cardsMined: number,
  linesSeen: number,
  tokensSeen: number,
  episodesStartedDelta: number,
  episodesCompletedDelta: number,
  startedAtMs: number,
  endedAtMs: number,
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
  endedAtMs: number,
): void {
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
    .run(session.sessionId, endedAtMs, Date.now(), Date.now());

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

  const activeMs = telemetry
    ? asPositiveNumber(telemetry.active_watched_ms, session.activeWatchedMs)
    : session.activeWatchedMs;
  const cardsMined = telemetry
    ? asPositiveNumber(telemetry.cards_mined, session.cardsMined)
    : session.cardsMined;
  const linesSeen = telemetry
    ? asPositiveNumber(telemetry.lines_seen, session.linesSeen)
    : session.linesSeen;
  const tokensSeen = telemetry
    ? asPositiveNumber(telemetry.tokens_seen, session.tokensSeen)
    : session.tokensSeen;
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

  const nowMs = Date.now();
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
    nowMs,
  );

  upsertLifetimeMedia(
    db,
    session.videoId,
    nowMs,
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
      nowMs,
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
  const rebuiltAtMs = Date.now();
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
  rebuiltAtMs = Date.now(),
): LifetimeRebuildSummary {
  return rebuildLifetimeSummariesInternal(db, rebuiltAtMs);
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
