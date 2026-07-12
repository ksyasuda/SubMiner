import type { LexiconResolver } from './merge-catalog';
import { selectAll, selectOne, type SqlRow, type SyncDb } from './driver';
import { insertRow, nowDbTimestamp, type SyncMergeSummary } from './shared';

const SESSION_COPY_COLUMNS = [
  'session_uuid',
  'started_at_ms',
  'ended_at_ms',
  'status',
  'locale_id',
  'target_lang_id',
  'difficulty_tier',
  'subtitle_mode',
  'ended_media_ms',
  'total_watched_ms',
  'active_watched_ms',
  'lines_seen',
  'tokens_seen',
  'cards_mined',
  'lookup_count',
  'lookup_hits',
  'yomitan_lookup_count',
  'pause_count',
  'pause_ms',
  'seek_forward_count',
  'seek_backward_count',
  'media_buffer_events',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

const TELEMETRY_COPY_COLUMNS = [
  'sample_ms',
  'total_watched_ms',
  'active_watched_ms',
  'lines_seen',
  'tokens_seen',
  'cards_mined',
  'lookup_count',
  'lookup_hits',
  'yomitan_lookup_count',
  'pause_count',
  'pause_ms',
  'seek_forward_count',
  'seek_backward_count',
  'media_buffer_events',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

const EVENT_COPY_COLUMNS = [
  'ts_ms',
  'event_type',
  'line_index',
  'segment_start_ms',
  'segment_end_ms',
  'tokens_delta',
  'cards_delta',
  'payload_json',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

const LINE_COPY_COLUMNS = [
  'line_index',
  'segment_start_ms',
  'segment_end_ms',
  'text',
  'secondary_text',
  'CREATED_DATE',
  'LAST_UPDATE_DATE',
] as const;

export interface SessionMergeResult {
  newSessionIds: number[];
}

export function mergeSessions(
  local: SyncDb,
  remote: SyncDb,
  videoIdMap: Map<number, number>,
  animeIdMap: Map<number, number>,
  lexicon: LexiconResolver,
  summary: SyncMergeSummary,
): SessionMergeResult {
  const newSessionIds: number[] = [];
  const uuidExists = local.query('SELECT session_id FROM imm_sessions WHERE session_uuid = ?');

  const remoteSessions = selectAll(
    remote,
    `SELECT session_id, video_id, ${SESSION_COPY_COLUMNS.join(', ')}
     FROM imm_sessions
     ORDER BY CAST(started_at_ms AS REAL) ASC, session_id ASC`,
  );

  for (const session of remoteSessions) {
    if (session.ended_at_ms === null) {
      // Stale ACTIVE sessions are finalized by the app on its next startup;
      // they will sync once they carry final numbers.
      summary.activeSessionsSkipped += 1;
      continue;
    }
    if (uuidExists.get(session.session_uuid)) {
      summary.sessionsAlreadyPresent += 1;
      continue;
    }
    const localVideoId = videoIdMap.get(Number(session.video_id));
    if (localVideoId === undefined) {
      throw new Error(
        `Snapshot session ${String(session.session_uuid)} references missing video row`,
      );
    }

    const localSessionId = insertRow(
      local,
      'imm_sessions',
      ['video_id', ...SESSION_COPY_COLUMNS],
      [localVideoId, ...SESSION_COPY_COLUMNS.map((column) => session[column])],
    );
    newSessionIds.push(localSessionId);
    summary.sessionsMerged += 1;

    const remoteSessionId = Number(session.session_id);
    copyTelemetry(local, remote, remoteSessionId, localSessionId, summary);
    const eventIdMap = copyEvents(local, remote, remoteSessionId, localSessionId, summary);
    copySubtitleLines(
      local,
      remote,
      remoteSessionId,
      localSessionId,
      localVideoId,
      animeIdMap,
      eventIdMap,
      lexicon,
      summary,
    );
    applyMergedSessionLifetime(local, localSessionId, localVideoId, session);
  }

  return { newSessionIds };
}

function copyTelemetry(
  local: SyncDb,
  remote: SyncDb,
  remoteSessionId: number,
  localSessionId: number,
  summary: SyncMergeSummary,
): void {
  const rows = selectAll(
    remote,
    `SELECT ${TELEMETRY_COPY_COLUMNS.join(', ')} FROM imm_session_telemetry
     WHERE session_id = ? ORDER BY telemetry_id ASC`,
    [remoteSessionId],
  );
  for (const row of rows) {
    insertRow(
      local,
      'imm_session_telemetry',
      ['session_id', ...TELEMETRY_COPY_COLUMNS],
      [localSessionId, ...TELEMETRY_COPY_COLUMNS.map((column) => row[column])],
    );
    summary.telemetryRowsAdded += 1;
  }
}

function copyEvents(
  local: SyncDb,
  remote: SyncDb,
  remoteSessionId: number,
  localSessionId: number,
  summary: SyncMergeSummary,
): Map<number, number> {
  const eventIdMap = new Map<number, number>();
  const rows = selectAll(
    remote,
    `SELECT event_id, ${EVENT_COPY_COLUMNS.join(', ')} FROM imm_session_events
     WHERE session_id = ? ORDER BY event_id ASC`,
    [remoteSessionId],
  );
  for (const row of rows) {
    const localEventId = insertRow(
      local,
      'imm_session_events',
      ['session_id', ...EVENT_COPY_COLUMNS],
      [localSessionId, ...EVENT_COPY_COLUMNS.map((column) => row[column])],
    );
    eventIdMap.set(Number(row.event_id), localEventId);
    summary.eventsAdded += 1;
  }
  return eventIdMap;
}

function copySubtitleLines(
  local: SyncDb,
  remote: SyncDb,
  remoteSessionId: number,
  localSessionId: number,
  localVideoId: number,
  animeIdMap: Map<number, number>,
  eventIdMap: Map<number, number>,
  lexicon: LexiconResolver,
  summary: SyncMergeSummary,
): void {
  const rows = selectAll(
    remote,
    `SELECT line_id, event_id, anime_id, ${LINE_COPY_COLUMNS.join(', ')} FROM imm_subtitle_lines
     WHERE session_id = ? ORDER BY line_id ASC`,
    [remoteSessionId],
  );
  const wordOccurrences = remote.query(
    'SELECT word_id, occurrence_count FROM imm_word_line_occurrences WHERE line_id = ?',
  );
  const kanjiOccurrences = remote.query(
    'SELECT kanji_id, occurrence_count FROM imm_kanji_line_occurrences WHERE line_id = ?',
  );
  const insertWordOccurrence = local.query(
    `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count) VALUES (?, ?, ?)
     ON CONFLICT(line_id, word_id) DO UPDATE SET occurrence_count = occurrence_count + excluded.occurrence_count`,
  );
  const insertKanjiOccurrence = local.query(
    `INSERT INTO imm_kanji_line_occurrences (line_id, kanji_id, occurrence_count) VALUES (?, ?, ?)
     ON CONFLICT(line_id, kanji_id) DO UPDATE SET occurrence_count = occurrence_count + excluded.occurrence_count`,
  );

  for (const row of rows) {
    const localEventId =
      row.event_id === null ? null : (eventIdMap.get(Number(row.event_id)) ?? null);
    const localAnimeId =
      row.anime_id === null ? null : (animeIdMap.get(Number(row.anime_id)) ?? null);
    const localLineId = insertRow(
      local,
      'imm_subtitle_lines',
      ['session_id', 'event_id', 'video_id', 'anime_id', ...LINE_COPY_COLUMNS],
      [
        localSessionId,
        localEventId,
        localVideoId,
        localAnimeId,
        ...LINE_COPY_COLUMNS.map((column) => row[column]),
      ],
    );
    summary.subtitleLinesAdded += 1;

    for (const occurrence of wordOccurrences.all(row.line_id) as SqlRow[]) {
      const localWordId = lexicon.resolveWord(Number(occurrence.word_id));
      const count = Number(occurrence.occurrence_count);
      insertWordOccurrence.run(localLineId, localWordId, count);
      lexicon.addWordOccurrences(Number(occurrence.word_id), count);
    }
    for (const occurrence of kanjiOccurrences.all(row.line_id) as SqlRow[]) {
      const localKanjiId = lexicon.resolveKanji(Number(occurrence.kanji_id));
      const count = Number(occurrence.occurrence_count);
      insertKanjiOccurrence.run(localLineId, localKanjiId, count);
      lexicon.addKanjiOccurrences(Number(occurrence.kanji_id), count);
    }
  }
}

/**
 * Port of applySessionLifetimeSummary (src/core/services/immersion-tracker/
 * lifetime.ts) for sessions arriving out of chronological order. The
 * "first session of the day / for this video" checks are order-independent
 * here (any other session counts, not just earlier ones): the local machine
 * already credited active_days/episodes_started when its own session was
 * applied, even if the merged session started earlier that day.
 */
function applyMergedSessionLifetime(
  local: SyncDb,
  sessionId: number,
  videoId: number,
  session: SqlRow,
): void {
  const updatedAtMs = nowDbTimestamp();
  const applied = local
    .query(
      `INSERT INTO imm_lifetime_applied_sessions (session_id, applied_at_ms, CREATED_DATE, LAST_UPDATE_DATE)
       VALUES (?, ?, ?, ?)
       ON CONFLICT(session_id) DO NOTHING`,
    )
    .run(sessionId, session.ended_at_ms, updatedAtMs, updatedAtMs);
  if (applied.changes <= 0) return;

  const telemetry = selectOne(
    local,
    `SELECT active_watched_ms, cards_mined, lines_seen, tokens_seen
     FROM imm_session_telemetry
     WHERE session_id = ?
     ORDER BY sample_ms DESC, telemetry_id DESC
     LIMIT 1`,
    [sessionId],
  );

  const metric = (telemetryValue: unknown, sessionValue: unknown): number => {
    const fromTelemetry = telemetry ? Number(telemetryValue) : Number.NaN;
    const value = Number.isFinite(fromTelemetry) ? fromTelemetry : Number(sessionValue);
    return Math.max(0, Math.floor(Number.isFinite(value) ? value : 0));
  };
  const activeMs = metric(telemetry?.active_watched_ms, session.active_watched_ms);
  const cardsMined = metric(telemetry?.cards_mined, session.cards_mined);
  const linesSeen = metric(telemetry?.lines_seen, session.lines_seen);
  const tokensSeen = metric(telemetry?.tokens_seen, session.tokens_seen);

  const video = selectOne(local, 'SELECT anime_id, watched FROM imm_videos WHERE video_id = ?', [
    videoId,
  ]);
  const watched = Number(video?.watched ?? 0);
  const animeId =
    video?.anime_id === null || video?.anime_id === undefined ? null : Number(video.anime_id);

  const mediaLifetime = selectOne(
    local,
    'SELECT completed FROM imm_lifetime_media WHERE video_id = ?',
    [videoId],
  );
  const hasOtherSessionForVideo = Boolean(
    local
      .query('SELECT 1 FROM imm_sessions WHERE video_id = ? AND session_id != ? LIMIT 1')
      .get(videoId, sessionId),
  );
  const isFirstSessionForVideoRun = !mediaLifetime && !hasOtherSessionForVideo;
  const isFirstCompletedSessionForVideoRun = watched > 0 && Number(mediaLifetime?.completed ?? 0) <= 0;

  const hasOtherSessionOnDay = Boolean(
    local
      .query(
        `SELECT 1 FROM imm_sessions
         WHERE session_id != ?
           AND CAST(julianday(CAST(started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER)
             = CAST(julianday(CAST(? AS REAL) / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER)
         LIMIT 1`,
      )
      .get(sessionId, session.started_at_ms),
  );

  let animeCompletedDelta = 0;
  if (animeId !== null && watched > 0 && isFirstCompletedSessionForVideoRun) {
    const animeLifetime = selectOne(
      local,
      'SELECT episodes_completed FROM imm_lifetime_anime WHERE anime_id = ?',
      [animeId],
    );
    const anime = selectOne(local, 'SELECT episodes_total FROM imm_anime WHERE anime_id = ?', [
      animeId,
    ]);
    const episodesCompletedBefore = Number(animeLifetime?.episodes_completed ?? 0);
    const episodesTotal =
      anime?.episodes_total === null || anime?.episodes_total === undefined
        ? null
        : Number(anime.episodes_total);
    if (
      episodesTotal !== null &&
      episodesTotal > 0 &&
      episodesCompletedBefore < episodesTotal &&
      episodesCompletedBefore + 1 >= episodesTotal
    ) {
      animeCompletedDelta = 1;
    }
  }

  local
    .query(
      `UPDATE imm_lifetime_global
       SET total_sessions = total_sessions + 1,
           total_active_ms = total_active_ms + ?,
           total_cards = total_cards + ?,
           active_days = active_days + ?,
           episodes_started = episodes_started + ?,
           episodes_completed = episodes_completed + ?,
           anime_completed = anime_completed + ?,
           LAST_UPDATE_DATE = ?
       WHERE global_id = 1`,
    )
    .run(
      activeMs,
      cardsMined,
      hasOtherSessionOnDay ? 0 : 1,
      isFirstSessionForVideoRun ? 1 : 0,
      isFirstCompletedSessionForVideoRun ? 1 : 0,
      animeCompletedDelta,
      updatedAtMs,
    );

  local
    .query(
      `INSERT INTO imm_lifetime_media(
         video_id, total_sessions, total_active_ms, total_cards, total_lines_seen,
         total_tokens_seen, completed, first_watched_ms, last_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
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
         LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE`,
    )
    .run(
      videoId,
      activeMs,
      cardsMined,
      linesSeen,
      tokensSeen,
      watched > 0 ? 1 : 0,
      session.started_at_ms,
      session.ended_at_ms,
      updatedAtMs,
      updatedAtMs,
    );

  if (animeId !== null) {
    local
      .query(
        `INSERT INTO imm_lifetime_anime(
           anime_id, total_sessions, total_active_ms, total_cards, total_lines_seen,
           total_tokens_seen, episodes_started, episodes_completed, first_watched_ms,
           last_watched_ms, CREATED_DATE, LAST_UPDATE_DATE
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
           LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE`,
      )
      .run(
        animeId,
        activeMs,
        cardsMined,
        linesSeen,
        tokensSeen,
        isFirstSessionForVideoRun ? 1 : 0,
        isFirstCompletedSessionForVideoRun ? 1 : 0,
        session.started_at_ms,
        session.ended_at_ms,
        updatedAtMs,
        updatedAtMs,
      );
  }
}
