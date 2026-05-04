import type { DatabaseSync } from './sqlite';
import type {
  AnimeAnilistEntryRow,
  AnimeDetailRow,
  AnimeEpisodeRow,
  AnimeLibraryRow,
  AnimeWordRow,
  EpisodeCardEventRow,
  EpisodesPerDayRow,
  ImmersionSessionRollupRow,
  MediaArtRow,
  MediaDetailRow,
  MediaLibraryRow,
  NewAnimePerDayRow,
  SessionSummaryQueryRow,
  StreakCalendarRow,
  WatchTimePerAnimeRow,
} from './types';
import {
  ACTIVE_SESSION_METRICS_CTE,
  SESSION_WORD_COUNTS_CTE,
  SESSION_WORD_COUNTS_SELECT,
  fromDbTimestamp,
  resolvedCoverBlobExpr,
  sessionDisplayWordsExpr,
  visibleWordSql,
} from './query-shared';

export function getAnimeLibrary(db: DatabaseSync): AnimeLibraryRow[] {
  const rows = db
    .prepare(
      `
    SELECT
      a.anime_id AS animeId,
      a.canonical_title AS canonicalTitle,
      a.anilist_id AS anilistId,
      COALESCE(lm.total_sessions, 0) AS totalSessions,
      COALESCE(lm.total_active_ms, 0) AS totalActiveMs,
      COALESCE(lm.total_cards, 0) AS totalCards,
      COALESCE(lm.total_tokens_seen, 0) AS totalTokensSeen,
      COUNT(DISTINCT v.video_id) AS episodeCount,
      a.episodes_total AS episodesTotal,
      COALESCE(lm.last_watched_ms, 0) AS lastWatchedMs
    FROM imm_anime a
    JOIN imm_lifetime_anime lm ON lm.anime_id = a.anime_id
    JOIN imm_videos v ON v.anime_id = a.anime_id
    GROUP BY a.anime_id
    ORDER BY totalActiveMs DESC, lm.last_watched_ms DESC, canonicalTitle ASC
  `,
    )
    .all() as Array<AnimeLibraryRow & { lastWatchedMs: number | string }>;
  return rows.map((row) => ({
    ...row,
    lastWatchedMs: fromDbTimestamp(row.lastWatchedMs) ?? 0,
  }));
}

export function getAnimeDetail(db: DatabaseSync, animeId: number): AnimeDetailRow | null {
  const row = db
    .prepare(
      `
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      a.anime_id AS animeId,
      a.canonical_title AS canonicalTitle,
      a.anilist_id AS anilistId,
      a.title_romaji AS titleRomaji,
      a.title_english AS titleEnglish,
      a.title_native AS titleNative,
      a.description AS description,
      COALESCE(lm.total_sessions, 0) AS totalSessions,
      COALESCE(lm.total_active_ms, 0) AS totalActiveMs,
      COALESCE(lm.total_cards, 0) AS totalCards,
      COALESCE(lm.total_tokens_seen, 0) AS totalTokensSeen,
      COALESCE(lm.total_lines_seen, 0) AS totalLinesSeen,
      COALESCE(SUM(COALESCE(asm.lookupCount, s.lookup_count, 0)), 0) AS totalLookupCount,
      COALESCE(SUM(COALESCE(asm.lookupHits, s.lookup_hits, 0)), 0) AS totalLookupHits,
      COALESCE(SUM(COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0)), 0) AS totalYomitanLookupCount,
      COUNT(DISTINCT v.video_id) AS episodeCount,
      COALESCE(lm.last_watched_ms, 0) AS lastWatchedMs
    FROM imm_anime a
    JOIN imm_lifetime_anime lm ON lm.anime_id = a.anime_id
    JOIN imm_videos v ON v.anime_id = a.anime_id
    LEFT JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    WHERE a.anime_id = ?
    GROUP BY a.anime_id
  `,
    )
    .get(animeId) as (AnimeDetailRow & { lastWatchedMs: number | string }) | null;
  return row
    ? {
        ...row,
        lastWatchedMs: fromDbTimestamp(row.lastWatchedMs) ?? 0,
      }
    : null;
}

export function getAnimeAnilistEntries(db: DatabaseSync, animeId: number): AnimeAnilistEntryRow[] {
  return db
    .prepare(
      `
    SELECT DISTINCT
      m.anilist_id AS anilistId,
      m.title_romaji AS titleRomaji,
      m.title_english AS titleEnglish,
      v.parsed_season AS season
    FROM imm_videos v
    JOIN imm_media_art m ON m.video_id = v.video_id
    WHERE v.anime_id = ?
      AND m.anilist_id IS NOT NULL
    ORDER BY v.parsed_season ASC
  `,
    )
    .all(animeId) as unknown as AnimeAnilistEntryRow[];
}

export function getAnimeEpisodes(db: DatabaseSync, animeId: number): AnimeEpisodeRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc', 'COALESCE(asm.tokensSeen, s.tokens_seen)');
  const rows = db
    .prepare(
      `
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      v.anime_id AS animeId,
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      v.parsed_title AS parsedTitle,
      v.parsed_season AS season,
      v.parsed_episode AS episode,
      v.duration_ms AS durationMs,
      (
        SELECT COALESCE(
          NULLIF(s_recent.ended_media_ms, 0),
          (
            SELECT MAX(line.segment_end_ms)
            FROM imm_subtitle_lines line
            WHERE line.session_id = s_recent.session_id
              AND line.segment_end_ms IS NOT NULL
          ),
          (
            SELECT MAX(event.segment_end_ms)
            FROM imm_session_events event
            WHERE event.session_id = s_recent.session_id
              AND event.segment_end_ms IS NOT NULL
          )
        )
        FROM imm_sessions s_recent
        WHERE s_recent.video_id = v.video_id
          AND (
            s_recent.ended_media_ms IS NOT NULL
            OR EXISTS (
              SELECT 1
              FROM imm_subtitle_lines line
              WHERE line.session_id = s_recent.session_id
                AND line.segment_end_ms IS NOT NULL
            )
            OR EXISTS (
              SELECT 1
              FROM imm_session_events event
              WHERE event.session_id = s_recent.session_id
                AND event.segment_end_ms IS NOT NULL
            )
          )
        ORDER BY
          COALESCE(s_recent.ended_at_ms, s_recent.LAST_UPDATE_DATE, s_recent.started_at_ms) DESC,
          s_recent.session_id DESC
        LIMIT 1
      ) AS endedMediaMs,
      v.watched AS watched,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0)), 0) AS totalActiveMs,
      COALESCE(SUM(COALESCE(asm.cardsMined, s.cards_mined, 0)), 0) AS totalCards,
      COALESCE(SUM(${wordsExpr}), 0) AS totalTokensSeen,
      COALESCE(SUM(COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0)), 0) AS totalYomitanLookupCount,
      MAX(s.started_at_ms) AS lastWatchedMs
    FROM imm_videos v
    LEFT JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN session_word_counts swc ON swc.sessionId = s.session_id
    WHERE v.anime_id = ?
    GROUP BY v.video_id
    ORDER BY
      CASE WHEN v.parsed_season IS NULL THEN 1 ELSE 0 END,
      v.parsed_season ASC,
      CASE WHEN v.parsed_episode IS NULL THEN 1 ELSE 0 END,
      v.parsed_episode ASC,
      v.video_id ASC
  `,
    )
    .all(animeId) as Array<
    AnimeEpisodeRow & {
      endedMediaMs: number | string | null;
      lastWatchedMs: number | string;
    }
  >;
  return rows.map((row) => ({
    ...row,
    endedMediaMs: fromDbTimestamp(row.endedMediaMs),
    lastWatchedMs: fromDbTimestamp(row.lastWatchedMs) ?? 0,
  }));
}

export function getMediaLibrary(db: DatabaseSync): MediaLibraryRow[] {
  const rows = db
    .prepare(
      `
    SELECT
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      COALESCE(lm.total_sessions, 0) AS totalSessions,
      COALESCE(lm.total_active_ms, 0) AS totalActiveMs,
      COALESCE(lm.total_cards, 0) AS totalCards,
      COALESCE(lm.total_tokens_seen, 0) AS totalTokensSeen,
      COALESCE(lm.last_watched_ms, 0) AS lastWatchedMs,
      yv.youtube_video_id AS youtubeVideoId,
      yv.video_url AS videoUrl,
      yv.video_title AS videoTitle,
      yv.video_thumbnail_url AS videoThumbnailUrl,
      yv.channel_id AS channelId,
      yv.channel_name AS channelName,
      yv.channel_url AS channelUrl,
      yv.channel_thumbnail_url AS channelThumbnailUrl,
      yv.uploader_id AS uploaderId,
      yv.uploader_url AS uploaderUrl,
      yv.description AS description,
      CASE
        WHEN ma.cover_blob_hash IS NOT NULL OR ma.cover_blob IS NOT NULL THEN 1
        ELSE 0
      END AS hasCoverArt
    FROM imm_videos v
    JOIN imm_lifetime_media lm ON lm.video_id = v.video_id
    LEFT JOIN imm_media_art ma ON ma.video_id = v.video_id
    LEFT JOIN imm_youtube_videos yv ON yv.video_id = v.video_id
    ORDER BY lm.last_watched_ms DESC
  `,
    )
    .all() as Array<MediaLibraryRow & { lastWatchedMs: number | string }>;
  return rows.map((row) => ({
    ...row,
    lastWatchedMs: fromDbTimestamp(row.lastWatchedMs) ?? 0,
  }));
}

export function getMediaDetail(db: DatabaseSync, videoId: number): MediaDetailRow | null {
  return db
    .prepare(
      `
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      v.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      v.anime_id AS animeId,
      COALESCE(lm.total_sessions, 0) AS totalSessions,
      COALESCE(lm.total_active_ms, 0) AS totalActiveMs,
      COALESCE(lm.total_cards, 0) AS totalCards,
      COALESCE(lm.total_tokens_seen, 0) AS totalTokensSeen,
      COALESCE(lm.total_lines_seen, 0) AS totalLinesSeen,
      COALESCE(SUM(COALESCE(asm.lookupCount, s.lookup_count, 0)), 0) AS totalLookupCount,
      COALESCE(SUM(COALESCE(asm.lookupHits, s.lookup_hits, 0)), 0) AS totalLookupHits,
      COALESCE(SUM(COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0)), 0) AS totalYomitanLookupCount,
      yv.youtube_video_id AS youtubeVideoId,
      yv.video_url AS videoUrl,
      yv.video_title AS videoTitle,
      yv.video_thumbnail_url AS videoThumbnailUrl,
      yv.channel_id AS channelId,
      yv.channel_name AS channelName,
      yv.channel_url AS channelUrl,
      yv.channel_thumbnail_url AS channelThumbnailUrl,
      yv.uploader_id AS uploaderId,
      yv.uploader_url AS uploaderUrl,
      yv.description AS description
    FROM imm_videos v
    JOIN imm_lifetime_media lm ON lm.video_id = v.video_id
    LEFT JOIN imm_youtube_videos yv ON yv.video_id = v.video_id
    LEFT JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    WHERE v.video_id = ?
    GROUP BY v.video_id
  `,
    )
    .get(videoId) as unknown as MediaDetailRow | null;
}

export function getMediaSessions(
  db: DatabaseSync,
  videoId: number,
  limit = 100,
): SessionSummaryQueryRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc', 'COALESCE(asm.tokensSeen, s.tokens_seen)');
  const rows = db
    .prepare(
      `
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      s.session_id AS sessionId,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      s.started_at_ms AS startedAtMs,
      s.ended_at_ms AS endedAtMs,
      COALESCE(asm.totalWatchedMs, s.total_watched_ms, 0) AS totalWatchedMs,
      COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0) AS activeWatchedMs,
      COALESCE(asm.linesSeen, s.lines_seen, 0) AS linesSeen,
      ${wordsExpr} AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.lookupCount, s.lookup_count, 0) AS lookupCount,
      COALESCE(asm.lookupHits, s.lookup_hits, 0) AS lookupHits,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount
    FROM imm_sessions s
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN session_word_counts swc ON swc.sessionId = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    WHERE s.video_id = ?
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `,
    )
    .all(videoId, limit) as Array<
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

export function getMediaDailyRollups(
  db: DatabaseSync,
  videoId: number,
  limit = 90,
): ImmersionSessionRollupRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc');
  return db
    .prepare(
      `
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
      WHERE video_id = ?
      ORDER BY rollup_day DESC
      LIMIT ?
    )
    SELECT
      rollup_day AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      CASE
        WHEN dwc.totalTokensSeen IS NOT NULL AND dwc.totalTokensSeen > total_tokens_seen THEN dwc.totalTokensSeen
        ELSE total_tokens_seen
      END AS totalTokensSeen,
      total_cards AS totalCards,
      cards_per_hour AS cardsPerHour,
      CASE
        WHEN total_active_min > 0 THEN (
          CASE
            WHEN dwc.totalTokensSeen IS NOT NULL AND dwc.totalTokensSeen > total_tokens_seen THEN dwc.totalTokensSeen
            ELSE total_tokens_seen
          END
        ) * 1.0 / total_active_min
        ELSE NULL
      END AS tokensPerMin,
      lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups
    LEFT JOIN daily_word_counts dwc
      ON dwc.rollupDay = rollup_day
      AND dwc.videoId = video_id
    WHERE video_id = ?
    AND rollup_day IN (SELECT rollup_day FROM recent_days)
    ORDER BY rollup_day DESC, video_id DESC
    `,
    )
    .all(videoId, limit, videoId) as unknown as ImmersionSessionRollupRow[];
}

export function getAnimeDailyRollups(
  db: DatabaseSync,
  animeId: number,
  limit = 90,
): ImmersionSessionRollupRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc');
  return db
    .prepare(
      `
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
      SELECT DISTINCT r.rollup_day
      FROM imm_daily_rollups r
      JOIN imm_videos v ON v.video_id = r.video_id
      WHERE v.anime_id = ?
      ORDER BY r.rollup_day DESC
      LIMIT ?
    )
    SELECT r.rollup_day AS rollupDayOrMonth, r.video_id AS videoId,
           r.total_sessions AS totalSessions, r.total_active_min AS totalActiveMin,
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
    JOIN imm_videos v ON v.video_id = r.video_id
    LEFT JOIN daily_word_counts dwc
      ON dwc.rollupDay = r.rollup_day
      AND dwc.videoId = r.video_id
    WHERE v.anime_id = ?
    AND r.rollup_day IN (SELECT rollup_day FROM recent_days)
    ORDER BY r.rollup_day DESC, r.video_id DESC
    `,
    )
    .all(animeId, limit, animeId) as unknown as ImmersionSessionRollupRow[];
}

export function getAnimeCoverArt(db: DatabaseSync, animeId: number): MediaArtRow | null {
  const resolvedCoverBlob = resolvedCoverBlobExpr('a', 'cab');
  const row = db
    .prepare(
      `
    SELECT
      a.video_id AS videoId,
      a.anilist_id AS anilistId,
      a.cover_url AS coverUrl,
      ${resolvedCoverBlob} AS coverBlob,
      a.title_romaji AS titleRomaji,
      a.title_english AS titleEnglish,
      a.episodes_total AS episodesTotal,
      a.fetched_at_ms AS fetchedAtMs
    FROM imm_media_art a
    JOIN imm_videos v ON v.video_id = a.video_id
    LEFT JOIN imm_cover_art_blobs cab ON cab.blob_hash = a.cover_blob_hash
    WHERE v.anime_id = ?
    AND ${resolvedCoverBlob} IS NOT NULL
    ORDER BY a.fetched_at_ms DESC, a.video_id DESC
    LIMIT 1
  `,
    )
    .get(animeId) as (MediaArtRow & { fetchedAtMs: number | string }) | null;
  return row
    ? {
        ...row,
        fetchedAtMs: fromDbTimestamp(row.fetchedAtMs) ?? 0,
      }
    : null;
}

export function getCoverArt(db: DatabaseSync, videoId: number): MediaArtRow | null {
  const resolvedCoverBlob = resolvedCoverBlobExpr('a', 'cab');
  const row = db
    .prepare(
      `
    SELECT
      a.video_id AS videoId,
      a.anilist_id AS anilistId,
      a.cover_url AS coverUrl,
      ${resolvedCoverBlob} AS coverBlob,
      a.title_romaji AS titleRomaji,
      a.title_english AS titleEnglish,
      a.episodes_total AS episodesTotal,
      a.fetched_at_ms AS fetchedAtMs
    FROM imm_media_art a
    LEFT JOIN imm_cover_art_blobs cab ON cab.blob_hash = a.cover_blob_hash
    WHERE a.video_id = ?
  `,
    )
    .get(videoId) as (MediaArtRow & { fetchedAtMs: number | string }) | null;
  return row
    ? {
        ...row,
        fetchedAtMs: fromDbTimestamp(row.fetchedAtMs) ?? 0,
      }
    : null;
}

export function getStreakCalendar(db: DatabaseSync, days = 90): StreakCalendarRow[] {
  const now = new Date();
  const localMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const todayLocalDay = Math.floor(localMidnight / 86_400_000);
  const cutoffDay = todayLocalDay - days;
  return db
    .prepare(
      `
    SELECT rollup_day AS epochDay, SUM(total_active_min) AS totalActiveMin
    FROM imm_daily_rollups
    WHERE rollup_day >= ?
    GROUP BY rollup_day
    ORDER BY rollup_day ASC
  `,
    )
    .all(cutoffDay) as StreakCalendarRow[];
}

export function getAnimeWords(db: DatabaseSync, animeId: number, limit = 50): AnimeWordRow[] {
  return db
    .prepare(
      `
    SELECT w.id AS wordId, w.headword, w.word, w.reading, w.part_of_speech AS partOfSpeech,
           SUM(o.occurrence_count) AS frequency
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_words w ON w.id = o.word_id
    WHERE sl.anime_id = ? AND ${visibleWordSql('w')}
    GROUP BY w.id
    ORDER BY frequency DESC
    LIMIT ?
  `,
    )
    .all(animeId, limit) as unknown as AnimeWordRow[];
}

export function getEpisodesPerDay(db: DatabaseSync, limit = 90): EpisodesPerDayRow[] {
  return db
    .prepare(
      `
    SELECT CAST(julianday(s.started_at_ms / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) AS epochDay,
           COUNT(DISTINCT s.video_id) AS episodeCount
    FROM imm_sessions s
    GROUP BY epochDay
    ORDER BY epochDay DESC
    LIMIT ?
  `,
    )
    .all(limit) as EpisodesPerDayRow[];
}

export function getNewAnimePerDay(db: DatabaseSync, limit = 90): NewAnimePerDayRow[] {
  return db
    .prepare(
      `
    SELECT first_day AS epochDay, COUNT(*) AS newAnimeCount
    FROM (
      SELECT CAST(julianday(MIN(s.started_at_ms) / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) AS first_day
      FROM imm_sessions s
      JOIN imm_videos v ON v.video_id = s.video_id
      WHERE v.anime_id IS NOT NULL
      GROUP BY v.anime_id
    )
    GROUP BY first_day
    ORDER BY first_day DESC
    LIMIT ?
  `,
    )
    .all(limit) as NewAnimePerDayRow[];
}

export function getWatchTimePerAnime(db: DatabaseSync, limit = 90): WatchTimePerAnimeRow[] {
  const nowD = new Date();
  const cutoffDay =
    Math.floor(
      new Date(nowD.getFullYear(), nowD.getMonth(), nowD.getDate()).getTime() / 86_400_000,
    ) - limit;
  return db
    .prepare(
      `
    SELECT r.rollup_day AS epochDay, a.anime_id AS animeId,
           a.canonical_title AS animeTitle,
           SUM(r.total_active_min) AS totalActiveMin
    FROM imm_daily_rollups r
    JOIN imm_videos v ON v.video_id = r.video_id
    JOIN imm_anime a ON a.anime_id = v.anime_id
    WHERE r.rollup_day >= ?
    GROUP BY r.rollup_day, a.anime_id
    ORDER BY r.rollup_day ASC
  `,
    )
    .all(cutoffDay) as WatchTimePerAnimeRow[];
}

export function getEpisodeWords(db: DatabaseSync, videoId: number, limit = 50): AnimeWordRow[] {
  return db
    .prepare(
      `
    SELECT w.id AS wordId, w.headword, w.word, w.reading, w.part_of_speech AS partOfSpeech,
           SUM(o.occurrence_count) AS frequency
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_words w ON w.id = o.word_id
    WHERE sl.video_id = ?
    GROUP BY w.id
    ORDER BY frequency DESC
    LIMIT ?
  `,
    )
    .all(videoId, limit) as unknown as AnimeWordRow[];
}

export function getEpisodeSessions(db: DatabaseSync, videoId: number): SessionSummaryQueryRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc', 'COALESCE(asm.tokensSeen, s.tokens_seen)');
  const rows = db
    .prepare(
      `
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      s.session_id AS sessionId, s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      s.started_at_ms AS startedAtMs, s.ended_at_ms AS endedAtMs,
      COALESCE(asm.totalWatchedMs, s.total_watched_ms, 0) AS totalWatchedMs,
      COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0) AS activeWatchedMs,
      COALESCE(asm.linesSeen, s.lines_seen, 0) AS linesSeen,
      ${wordsExpr} AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.lookupCount, s.lookup_count, 0) AS lookupCount,
      COALESCE(asm.lookupHits, s.lookup_hits, 0) AS lookupHits,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount
    FROM imm_sessions s
    JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN session_word_counts swc ON swc.sessionId = s.session_id
    WHERE s.video_id = ?
    ORDER BY s.started_at_ms DESC
  `,
    )
    .all(videoId) as Array<
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

export function getEpisodeCardEvents(db: DatabaseSync, videoId: number): EpisodeCardEventRow[] {
  const rows = db
    .prepare(
      `
    SELECT e.event_id AS eventId, e.session_id AS sessionId,
           e.ts_ms AS tsMs, e.cards_delta AS cardsDelta,
           e.payload_json AS payloadJson
    FROM imm_session_events e
    JOIN imm_sessions s ON s.session_id = e.session_id
    WHERE s.video_id = ? AND e.event_type = 4
    ORDER BY CAST(e.ts_ms AS REAL) DESC
  `,
    )
    .all(videoId) as Array<{
    eventId: number;
    sessionId: number;
    tsMs: number | string;
    cardsDelta: number;
    payloadJson: string | null;
  }>;

  return rows.map((row) => {
    let noteIds: number[] = [];
    if (row.payloadJson) {
      try {
        const parsed = JSON.parse(row.payloadJson);
        if (Array.isArray(parsed.noteIds)) noteIds = parsed.noteIds;
      } catch {}
    }
    return {
      eventId: row.eventId,
      sessionId: row.sessionId,
      tsMs: fromDbTimestamp(row.tsMs) ?? 0,
      cardsDelta: row.cardsDelta,
      noteIds,
    };
  });
}
