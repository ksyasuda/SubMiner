import { createHash } from 'node:crypto';
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
  KanjiAnimeAppearanceRow,
  KanjiDetailRow,
  KanjiOccurrenceRow,
  KanjiStatsRow,
  KanjiWordRow,
  MediaArtRow,
  MediaDetailRow,
  MediaLibraryRow,
  NewAnimePerDayRow,
  SessionEventRow,
  SessionSummaryQueryRow,
  SessionTimelineRow,
  SimilarWordRow,
  StreakCalendarRow,
  VocabularyCleanupSummary,
  WatchTimePerAnimeRow,
  WordAnimeAppearanceRow,
  WordDetailRow,
  WordOccurrenceRow,
  VocabularyStatsRow,
} from './types';
import { buildCoverBlobReference, normalizeCoverBlobBytes } from './storage';
import { PartOfSpeech, type MergedToken } from '../../../types';
import { shouldExcludeTokenFromVocabularyPersistence } from '../tokenizer/annotation-stage';
import { deriveStoredPartOfSpeech } from '../tokenizer/part-of-speech';

type CleanupVocabularyRow = {
  id: number;
  word: string;
  headword: string;
  reading: string | null;
  part_of_speech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
  first_seen: number | null;
  last_seen: number | null;
  frequency: number | null;
};

type ResolvedVocabularyPos = {
  headword: string;
  reading: string;
  hasPosMetadata: boolean;
  partOfSpeech: PartOfSpeech;
  pos1: string;
  pos2: string;
  pos3: string;
};

type CleanupVocabularyStatsOptions = {
  resolveLegacyPos?: (row: CleanupVocabularyRow) => Promise<{
    headword: string;
    reading: string;
    partOfSpeech: string;
    pos1: string;
    pos2: string;
    pos3: string;
  } | null>;
};

const ACTIVE_SESSION_METRICS_CTE = `
  WITH active_session_metrics AS (
    SELECT
      t.session_id AS sessionId,
      MAX(t.total_watched_ms) AS totalWatchedMs,
      MAX(t.active_watched_ms) AS activeWatchedMs,
      MAX(t.lines_seen) AS linesSeen,
      MAX(t.tokens_seen) AS tokensSeen,
      MAX(t.cards_mined) AS cardsMined,
      MAX(t.lookup_count) AS lookupCount,
      MAX(t.lookup_hits) AS lookupHits,
      MAX(t.yomitan_lookup_count) AS yomitanLookupCount
    FROM imm_session_telemetry t
    JOIN imm_sessions s ON s.session_id = t.session_id
    WHERE s.ended_at_ms IS NULL
    GROUP BY t.session_id
  )
`;

function resolvedCoverBlobExpr(mediaAlias: string, blobStoreAlias: string): string {
  return `COALESCE(${blobStoreAlias}.cover_blob, CASE WHEN ${mediaAlias}.cover_blob_hash IS NULL THEN ${mediaAlias}.cover_blob ELSE NULL END)`;
}

function cleanupUnusedCoverArtBlobHash(db: DatabaseSync, blobHash: string | null): void {
  if (!blobHash) {
    return;
  }
  db.prepare(
    `
      DELETE FROM imm_cover_art_blobs
      WHERE blob_hash = ?
        AND NOT EXISTS (
          SELECT 1
          FROM imm_media_art
          WHERE cover_blob_hash = ?
        )
    `,
  ).run(blobHash, blobHash);
}

function findSharedCoverBlobHash(
  db: DatabaseSync,
  videoId: number,
  anilistId: number | null,
  coverUrl: string | null,
): string | null {
  if (anilistId !== null) {
    const byAnilist = db
      .prepare(
        `
          SELECT cover_blob_hash AS coverBlobHash
          FROM imm_media_art
          WHERE video_id != ?
            AND anilist_id = ?
            AND cover_blob_hash IS NOT NULL
          ORDER BY fetched_at_ms DESC, video_id DESC
          LIMIT 1
        `,
      )
      .get(videoId, anilistId) as { coverBlobHash: string | null } | undefined;
    if (byAnilist?.coverBlobHash) {
      return byAnilist.coverBlobHash;
    }
  }

  if (coverUrl) {
    const byUrl = db
      .prepare(
        `
          SELECT cover_blob_hash AS coverBlobHash
          FROM imm_media_art
          WHERE video_id != ?
            AND cover_url = ?
            AND cover_blob_hash IS NOT NULL
          ORDER BY fetched_at_ms DESC, video_id DESC
          LIMIT 1
        `,
      )
      .get(videoId, coverUrl) as { coverBlobHash: string | null } | undefined;
    return byUrl?.coverBlobHash ?? null;
  }

  return null;
}

function makePlaceholders(values: number[]): string {
  return values.map(() => '?').join(',');
}

function getAffectedWordIdsForSessions(db: DatabaseSync, sessionIds: number[]): number[] {
  if (sessionIds.length === 0) {
    return [];
  }

  return (
    db
      .prepare(
        `
          SELECT DISTINCT o.word_id AS wordId
          FROM imm_word_line_occurrences o
          JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
          WHERE sl.session_id IN (${makePlaceholders(sessionIds)})
        `,
      )
      .all(...sessionIds) as Array<{ wordId: number }>
  ).map((row) => row.wordId);
}

function getAffectedKanjiIdsForSessions(db: DatabaseSync, sessionIds: number[]): number[] {
  if (sessionIds.length === 0) {
    return [];
  }

  return (
    db
      .prepare(
        `
          SELECT DISTINCT o.kanji_id AS kanjiId
          FROM imm_kanji_line_occurrences o
          JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
          WHERE sl.session_id IN (${makePlaceholders(sessionIds)})
        `,
      )
      .all(...sessionIds) as Array<{ kanjiId: number }>
  ).map((row) => row.kanjiId);
}

function getAffectedWordIdsForVideo(db: DatabaseSync, videoId: number): number[] {
  return (
    db
      .prepare(
        `
          SELECT DISTINCT o.word_id AS wordId
          FROM imm_word_line_occurrences o
          JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
          WHERE sl.video_id = ?
        `,
      )
      .all(videoId) as Array<{ wordId: number }>
  ).map((row) => row.wordId);
}

function getAffectedKanjiIdsForVideo(db: DatabaseSync, videoId: number): number[] {
  return (
    db
      .prepare(
        `
          SELECT DISTINCT o.kanji_id AS kanjiId
          FROM imm_kanji_line_occurrences o
          JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
          WHERE sl.video_id = ?
        `,
      )
      .all(videoId) as Array<{ kanjiId: number }>
  ).map((row) => row.kanjiId);
}

function refreshWordAggregates(db: DatabaseSync, wordIds: number[]): void {
  if (wordIds.length === 0) {
    return;
  }

  const rows = db
    .prepare(
      `
        SELECT
          w.id AS wordId,
          COALESCE(SUM(o.occurrence_count), 0) AS frequency,
          MIN(COALESCE(sl.CREATED_DATE, sl.LAST_UPDATE_DATE)) AS firstSeen,
          MAX(COALESCE(sl.LAST_UPDATE_DATE, sl.CREATED_DATE)) AS lastSeen
        FROM imm_words w
        LEFT JOIN imm_word_line_occurrences o ON o.word_id = w.id
        LEFT JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE w.id IN (${makePlaceholders(wordIds)})
        GROUP BY w.id
      `,
    )
    .all(...wordIds) as Array<{
    wordId: number;
    frequency: number;
    firstSeen: number | null;
    lastSeen: number | null;
  }>;
  const updateStmt = db.prepare(
    `
      UPDATE imm_words
      SET frequency = ?, first_seen = ?, last_seen = ?
      WHERE id = ?
    `,
  );
  const deleteStmt = db.prepare('DELETE FROM imm_words WHERE id = ?');

  for (const row of rows) {
    if (row.frequency <= 0 || row.firstSeen === null || row.lastSeen === null) {
      deleteStmt.run(row.wordId);
      continue;
    }
    updateStmt.run(row.frequency, row.firstSeen, row.lastSeen, row.wordId);
  }
}

function refreshKanjiAggregates(db: DatabaseSync, kanjiIds: number[]): void {
  if (kanjiIds.length === 0) {
    return;
  }

  const rows = db
    .prepare(
      `
        SELECT
          k.id AS kanjiId,
          COALESCE(SUM(o.occurrence_count), 0) AS frequency,
          MIN(COALESCE(sl.CREATED_DATE, sl.LAST_UPDATE_DATE)) AS firstSeen,
          MAX(COALESCE(sl.LAST_UPDATE_DATE, sl.CREATED_DATE)) AS lastSeen
        FROM imm_kanji k
        LEFT JOIN imm_kanji_line_occurrences o ON o.kanji_id = k.id
        LEFT JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
        WHERE k.id IN (${makePlaceholders(kanjiIds)})
        GROUP BY k.id
      `,
    )
    .all(...kanjiIds) as Array<{
    kanjiId: number;
    frequency: number;
    firstSeen: number | null;
    lastSeen: number | null;
  }>;
  const updateStmt = db.prepare(
    `
      UPDATE imm_kanji
      SET frequency = ?, first_seen = ?, last_seen = ?
      WHERE id = ?
    `,
  );
  const deleteStmt = db.prepare('DELETE FROM imm_kanji WHERE id = ?');

  for (const row of rows) {
    if (row.frequency <= 0 || row.firstSeen === null || row.lastSeen === null) {
      deleteStmt.run(row.kanjiId);
      continue;
    }
    updateStmt.run(row.frequency, row.firstSeen, row.lastSeen, row.kanjiId);
  }
}

function refreshLexicalAggregates(db: DatabaseSync, wordIds: number[], kanjiIds: number[]): void {
  refreshWordAggregates(db, [...new Set(wordIds)]);
  refreshKanjiAggregates(db, [...new Set(kanjiIds)]);
}

function deleteSessionsByIds(db: DatabaseSync, sessionIds: number[]): void {
  if (sessionIds.length === 0) {
    return;
  }

  const placeholders = makePlaceholders(sessionIds);
  db.prepare(`DELETE FROM imm_subtitle_lines WHERE session_id IN (${placeholders})`).run(
    ...sessionIds,
  );
  db.prepare(`DELETE FROM imm_session_telemetry WHERE session_id IN (${placeholders})`).run(
    ...sessionIds,
  );
  db.prepare(`DELETE FROM imm_session_events WHERE session_id IN (${placeholders})`).run(
    ...sessionIds,
  );
  db.prepare(`DELETE FROM imm_sessions WHERE session_id IN (${placeholders})`).run(...sessionIds);
}

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
  return prepared.all(limit) as unknown as SessionSummaryQueryRow[];
}

export function getSessionTimeline(
  db: DatabaseSync,
  sessionId: number,
  limit?: number,
): SessionTimelineRow[] {
  if (limit === undefined) {
    const prepared = db.prepare(`
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
    `);
    return prepared.all(sessionId) as unknown as SessionTimelineRow[];
  }

  const prepared = db.prepare(`
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
    LIMIT ?
  `);
  return prepared.all(sessionId, limit) as unknown as SessionTimelineRow[];
}

/** Returns all distinct headwords in the vocabulary table (global). */
export function getAllDistinctHeadwords(db: DatabaseSync): string[] {
  const rows = db.prepare('SELECT DISTINCT headword FROM imm_words').all() as Array<{
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
      WHERE sl.anime_id = ?
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
      WHERE sl.video_id = ?
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
    WHERE sl.session_id = ?
    ORDER BY sl.line_index ASC
  `);
  return stmt.all(sessionId) as Array<{
    lineIndex: number;
    headword: string;
    occurrenceCount: number;
  }>;
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

  const now = new Date();
  const todayLocal = Math.floor(
    new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() / 86_400_000,
  );
  const episodesToday =
    (
      db
        .prepare(
          `
    SELECT COUNT(DISTINCT s.video_id) AS count
    FROM imm_sessions s
    WHERE CAST(julianday(s.started_at_ms / 1000, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) = ?
  `,
        )
        .get(todayLocal) as { count: number }
    )?.count ?? 0;

  const thirtyDaysAgoMs = Date.now() - 30 * 86400000;
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

  const lookupTotals = db
    .prepare(
      `
    SELECT
      COALESCE(SUM(COALESCE(t.tokens_seen, s.tokens_seen, 0)), 0) AS totalTokensSeen,
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

function getNewWordCounts(db: DatabaseSync): { newWordsToday: number; newWordsThisWeek: number } {
  const now = new Date();
  const todayStartSec = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() / 1000;
  const weekAgoSec = todayStartSec - 7 * 86_400;

  const row = db
    .prepare(
      `
    WITH headword_first_seen AS (
      SELECT
        headword,
        MIN(first_seen) AS first_seen
      FROM imm_words
      WHERE first_seen IS NOT NULL
        AND headword IS NOT NULL
        AND headword != ''
      GROUP BY headword
    )
    SELECT
      COALESCE(SUM(CASE WHEN first_seen >= ? THEN 1 ELSE 0 END), 0) AS today,
      COALESCE(SUM(CASE WHEN first_seen >= ? THEN 1 ELSE 0 END), 0) AS week
    FROM headword_first_seen
  `,
    )
    .get(todayStartSec, weekAgoSec) as { today: number; week: number } | null;

  return {
    newWordsToday: Number(row?.today ?? 0),
    newWordsThisWeek: Number(row?.week ?? 0),
  };
}

export function getDailyRollups(db: DatabaseSync, limit = 60): ImmersionSessionRollupRow[] {
  const prepared = db.prepare(`
    WITH recent_days AS (
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
      r.total_tokens_seen AS totalTokensSeen,
      r.total_cards AS totalCards,
      r.cards_per_hour AS cardsPerHour,
      r.tokens_per_min AS tokensPerMin,
      r.lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups r
    WHERE r.rollup_day IN (SELECT rollup_day FROM recent_days)
    ORDER BY r.rollup_day DESC, r.video_id DESC
    `);

  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}

export function getMonthlyRollups(db: DatabaseSync, limit = 24): ImmersionSessionRollupRow[] {
  const prepared = db.prepare(`
    WITH recent_months AS (
      SELECT DISTINCT rollup_month
      FROM imm_monthly_rollups
      ORDER BY rollup_month DESC
      LIMIT ?
    )
    SELECT
      rollup_month AS rollupDayOrMonth,
      video_id AS videoId,
      total_sessions AS totalSessions,
      total_active_min AS totalActiveMin,
      total_lines_seen AS totalLinesSeen,
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      0 AS cardsPerHour,
      0 AS tokensPerMin,
      0 AS lookupHitRate
    FROM imm_monthly_rollups
    WHERE rollup_month IN (SELECT rollup_month FROM recent_months)
    ORDER BY rollup_month DESC, video_id DESC
  `);
  return prepared.all(limit) as unknown as ImmersionSessionRollupRow[];
}

type TrendRange = '7d' | '30d' | '90d' | 'all';
type TrendGroupBy = 'day' | 'month';

interface TrendChartPoint {
  label: string;
  value: number;
}

interface TrendPerAnimePoint {
  epochDay: number;
  animeTitle: string;
  value: number;
}

interface TrendSessionMetricRow {
  startedAtMs: number;
  videoId: number | null;
  canonicalTitle: string | null;
  animeTitle: string | null;
  activeWatchedMs: number;
  tokensSeen: number;
  cardsMined: number;
  yomitanLookupCount: number;
}

export interface TrendsDashboardQueryResult {
  activity: {
    watchTime: TrendChartPoint[];
    cards: TrendChartPoint[];
    words: TrendChartPoint[];
    sessions: TrendChartPoint[];
  };
  progress: {
    watchTime: TrendChartPoint[];
    sessions: TrendChartPoint[];
    words: TrendChartPoint[];
    newWords: TrendChartPoint[];
    cards: TrendChartPoint[];
    episodes: TrendChartPoint[];
    lookups: TrendChartPoint[];
  };
  ratios: {
    lookupsPerHundred: TrendChartPoint[];
  };
  animePerDay: {
    episodes: TrendPerAnimePoint[];
    watchTime: TrendPerAnimePoint[];
    cards: TrendPerAnimePoint[];
    words: TrendPerAnimePoint[];
    lookups: TrendPerAnimePoint[];
    lookupsPerHundred: TrendPerAnimePoint[];
  };
  animeCumulative: {
    watchTime: TrendPerAnimePoint[];
    episodes: TrendPerAnimePoint[];
    cards: TrendPerAnimePoint[];
    words: TrendPerAnimePoint[];
  };
  patterns: {
    watchTimeByDayOfWeek: TrendChartPoint[];
    watchTimeByHour: TrendChartPoint[];
  };
}

const TREND_DAY_LIMITS: Record<Exclude<TrendRange, 'all'>, number> = {
  '7d': 7,
  '30d': 30,
  '90d': 90,
};

const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

function getTrendDayLimit(range: TrendRange): number {
  return range === 'all' ? 365 : TREND_DAY_LIMITS[range];
}

function getTrendMonthlyLimit(range: TrendRange): number {
  if (range === 'all') {
    return 120;
  }
  return Math.max(1, Math.ceil(TREND_DAY_LIMITS[range] / 30));
}

function getTrendCutoffMs(range: TrendRange): number | null {
  if (range === 'all') {
    return null;
  }
  const dayLimit = getTrendDayLimit(range);
  const now = new Date();
  const localMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  return localMidnight - (dayLimit - 1) * 86_400_000;
}

function makeTrendLabel(value: number): string {
  if (value > 100_000) {
    const year = Math.floor(value / 100);
    const month = value % 100;
    return new Date(Date.UTC(year, month - 1, 1)).toLocaleDateString(undefined, {
      month: 'short',
      year: '2-digit',
    });
  }

  return new Date(value * 86_400_000).toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
  });
}

function getTrendSessionWordCount(session: Pick<TrendSessionMetricRow, 'tokensSeen'>): number {
  return session.tokensSeen;
}

function resolveTrendAnimeTitle(value: {
  animeTitle: string | null;
  canonicalTitle: string | null;
}): string {
  return value.animeTitle ?? value.canonicalTitle ?? 'Unknown';
}

function accumulatePoints(points: TrendChartPoint[]): TrendChartPoint[] {
  let sum = 0;
  return points.map((point) => {
    sum += point.value;
    return {
      label: point.label,
      value: sum,
    };
  });
}

function buildAggregatedTrendRows(rollups: ImmersionSessionRollupRow[]) {
  const byKey = new Map<
    number,
    { activeMin: number; cards: number; words: number; sessions: number }
  >();

  for (const rollup of rollups) {
    const existing = byKey.get(rollup.rollupDayOrMonth) ?? {
      activeMin: 0,
      cards: 0,
      words: 0,
      sessions: 0,
    };
    existing.activeMin += Math.round(rollup.totalActiveMin);
    existing.cards += rollup.totalCards;
    existing.words += rollup.totalTokensSeen;
    existing.sessions += rollup.totalSessions;
    byKey.set(rollup.rollupDayOrMonth, existing);
  }

  return Array.from(byKey.entries())
    .sort(([left], [right]) => left - right)
    .map(([key, value]) => ({
      label: makeTrendLabel(key),
      activeMin: value.activeMin,
      cards: value.cards,
      words: value.words,
      sessions: value.sessions,
    }));
}

function buildWatchTimeByDayOfWeek(sessions: TrendSessionMetricRow[]): TrendChartPoint[] {
  const totals = new Array(7).fill(0);
  for (const session of sessions) {
    totals[new Date(session.startedAtMs).getDay()] += session.activeWatchedMs;
  }
  return DAY_NAMES.map((name, index) => ({
    label: name,
    value: Math.round(totals[index] / 60_000),
  }));
}

function buildWatchTimeByHour(sessions: TrendSessionMetricRow[]): TrendChartPoint[] {
  const totals = new Array(24).fill(0);
  for (const session of sessions) {
    totals[new Date(session.startedAtMs).getHours()] += session.activeWatchedMs;
  }
  return totals.map((ms, index) => ({
    label: `${String(index).padStart(2, '0')}:00`,
    value: Math.round(ms / 60_000),
  }));
}

function dayLabel(epochDay: number): string {
  return new Date(epochDay * 86_400_000).toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
  });
}

function buildSessionSeriesByDay(
  sessions: TrendSessionMetricRow[],
  getValue: (session: TrendSessionMetricRow) => number,
): TrendChartPoint[] {
  const byDay = new Map<number, number>();
  for (const session of sessions) {
    const epochDay = Math.floor(session.startedAtMs / 86_400_000);
    byDay.set(epochDay, (byDay.get(epochDay) ?? 0) + getValue(session));
  }
  return Array.from(byDay.entries())
    .sort(([left], [right]) => left - right)
    .map(([epochDay, value]) => ({ label: dayLabel(epochDay), value }));
}

function buildLookupsPerHundredWords(sessions: TrendSessionMetricRow[]): TrendChartPoint[] {
  const lookupsByDay = new Map<number, number>();
  const wordsByDay = new Map<number, number>();

  for (const session of sessions) {
    const epochDay = Math.floor(session.startedAtMs / 86_400_000);
    lookupsByDay.set(epochDay, (lookupsByDay.get(epochDay) ?? 0) + session.yomitanLookupCount);
    wordsByDay.set(epochDay, (wordsByDay.get(epochDay) ?? 0) + getTrendSessionWordCount(session));
  }

  return Array.from(lookupsByDay.entries())
    .sort(([left], [right]) => left - right)
    .map(([epochDay, lookups]) => {
      const words = wordsByDay.get(epochDay) ?? 0;
      return {
        label: dayLabel(epochDay),
        value: words > 0 ? +((lookups / words) * 100).toFixed(1) : 0,
      };
    });
}

function buildPerAnimeFromSessions(
  sessions: TrendSessionMetricRow[],
  getValue: (session: TrendSessionMetricRow) => number,
): TrendPerAnimePoint[] {
  const byAnime = new Map<string, Map<number, number>>();

  for (const session of sessions) {
    const animeTitle = resolveTrendAnimeTitle(session);
    const epochDay = Math.floor(session.startedAtMs / 86_400_000);
    const dayMap = byAnime.get(animeTitle) ?? new Map();
    dayMap.set(epochDay, (dayMap.get(epochDay) ?? 0) + getValue(session));
    byAnime.set(animeTitle, dayMap);
  }

  const result: TrendPerAnimePoint[] = [];
  for (const [animeTitle, dayMap] of byAnime) {
    for (const [epochDay, value] of dayMap) {
      result.push({ epochDay, animeTitle, value });
    }
  }
  return result;
}

function buildLookupsPerHundredPerAnime(sessions: TrendSessionMetricRow[]): TrendPerAnimePoint[] {
  const lookups = new Map<string, Map<number, number>>();
  const words = new Map<string, Map<number, number>>();

  for (const session of sessions) {
    const animeTitle = resolveTrendAnimeTitle(session);
    const epochDay = Math.floor(session.startedAtMs / 86_400_000);

    const lookupMap = lookups.get(animeTitle) ?? new Map();
    lookupMap.set(epochDay, (lookupMap.get(epochDay) ?? 0) + session.yomitanLookupCount);
    lookups.set(animeTitle, lookupMap);

    const wordMap = words.get(animeTitle) ?? new Map();
    wordMap.set(epochDay, (wordMap.get(epochDay) ?? 0) + getTrendSessionWordCount(session));
    words.set(animeTitle, wordMap);
  }

  const result: TrendPerAnimePoint[] = [];
  for (const [animeTitle, dayMap] of lookups) {
    const wordMap = words.get(animeTitle) ?? new Map();
    for (const [epochDay, lookupCount] of dayMap) {
      const wordCount = wordMap.get(epochDay) ?? 0;
      result.push({
        epochDay,
        animeTitle,
        value: wordCount > 0 ? +((lookupCount / wordCount) * 100).toFixed(1) : 0,
      });
    }
  }
  return result;
}

function buildCumulativePerAnime(points: TrendPerAnimePoint[]): TrendPerAnimePoint[] {
  const byAnime = new Map<string, Map<number, number>>();
  const allDays = new Set<number>();

  for (const point of points) {
    const dayMap = byAnime.get(point.animeTitle) ?? new Map();
    dayMap.set(point.epochDay, (dayMap.get(point.epochDay) ?? 0) + point.value);
    byAnime.set(point.animeTitle, dayMap);
    allDays.add(point.epochDay);
  }

  const sortedDays = [...allDays].sort((left, right) => left - right);
  if (sortedDays.length === 0) {
    return [];
  }

  const minDay = sortedDays[0]!;
  const maxDay = sortedDays[sortedDays.length - 1]!;
  const result: TrendPerAnimePoint[] = [];

  for (const [animeTitle, dayMap] of byAnime) {
    const firstDay = Math.min(...dayMap.keys());
    let cumulative = 0;
    for (let epochDay = minDay; epochDay <= maxDay; epochDay += 1) {
      if (epochDay < firstDay) {
        continue;
      }
      cumulative += dayMap.get(epochDay) ?? 0;
      result.push({ epochDay, animeTitle, value: cumulative });
    }
  }

  return result;
}

function getVideoAnimeTitleMap(
  db: DatabaseSync,
  videoIds: Array<number | null>,
): Map<number, string> {
  const uniqueIds = [
    ...new Set(videoIds.filter((value): value is number => typeof value === 'number')),
  ];
  if (uniqueIds.length === 0) {
    return new Map();
  }

  const rows = db
    .prepare(
      `
        SELECT
          v.video_id AS videoId,
          COALESCE(a.canonical_title, v.canonical_title, 'Unknown') AS animeTitle
        FROM imm_videos v
        LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
        WHERE v.video_id IN (${makePlaceholders(uniqueIds)})
      `,
    )
    .all(...uniqueIds) as Array<{ videoId: number; animeTitle: string }>;

  return new Map(rows.map((row) => [row.videoId, row.animeTitle]));
}

function resolveVideoAnimeTitle(
  videoId: number | null,
  titlesByVideoId: Map<number, string>,
): string {
  if (videoId === null) {
    return 'Unknown';
  }
  return titlesByVideoId.get(videoId) ?? 'Unknown';
}

function buildPerAnimeFromDailyRollups(
  rollups: ImmersionSessionRollupRow[],
  titlesByVideoId: Map<number, string>,
  getValue: (rollup: ImmersionSessionRollupRow) => number,
): TrendPerAnimePoint[] {
  const byAnime = new Map<string, Map<number, number>>();

  for (const rollup of rollups) {
    const animeTitle = resolveVideoAnimeTitle(rollup.videoId, titlesByVideoId);
    const dayMap = byAnime.get(animeTitle) ?? new Map();
    dayMap.set(
      rollup.rollupDayOrMonth,
      (dayMap.get(rollup.rollupDayOrMonth) ?? 0) + getValue(rollup),
    );
    byAnime.set(animeTitle, dayMap);
  }

  const result: TrendPerAnimePoint[] = [];
  for (const [animeTitle, dayMap] of byAnime) {
    for (const [epochDay, value] of dayMap) {
      result.push({ epochDay, animeTitle, value });
    }
  }
  return result;
}

function buildEpisodesPerAnimeFromDailyRollups(
  rollups: ImmersionSessionRollupRow[],
  titlesByVideoId: Map<number, string>,
): TrendPerAnimePoint[] {
  const byAnime = new Map<string, Map<number, Set<number>>>();

  for (const rollup of rollups) {
    if (rollup.videoId === null) {
      continue;
    }
    const animeTitle = resolveVideoAnimeTitle(rollup.videoId, titlesByVideoId);
    const dayMap = byAnime.get(animeTitle) ?? new Map();
    const videoIds = dayMap.get(rollup.rollupDayOrMonth) ?? new Set<number>();
    videoIds.add(rollup.videoId);
    dayMap.set(rollup.rollupDayOrMonth, videoIds);
    byAnime.set(animeTitle, dayMap);
  }

  const result: TrendPerAnimePoint[] = [];
  for (const [animeTitle, dayMap] of byAnime) {
    for (const [epochDay, videoIds] of dayMap) {
      result.push({ epochDay, animeTitle, value: videoIds.size });
    }
  }
  return result;
}

function buildEpisodesPerDayFromDailyRollups(
  rollups: ImmersionSessionRollupRow[],
): TrendChartPoint[] {
  const byDay = new Map<number, Set<number>>();

  for (const rollup of rollups) {
    if (rollup.videoId === null) {
      continue;
    }
    const videoIds = byDay.get(rollup.rollupDayOrMonth) ?? new Set<number>();
    videoIds.add(rollup.videoId);
    byDay.set(rollup.rollupDayOrMonth, videoIds);
  }

  return Array.from(byDay.entries())
    .sort(([left], [right]) => left - right)
    .map(([epochDay, videoIds]) => ({
      label: dayLabel(epochDay),
      value: videoIds.size,
    }));
}

function getTrendSessionMetrics(
  db: DatabaseSync,
  cutoffMs: number | null,
): TrendSessionMetricRow[] {
  const whereClause = cutoffMs === null ? '' : 'WHERE s.started_at_ms >= ?';
  const prepared = db.prepare(`
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      s.started_at_ms AS startedAtMs,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      a.canonical_title AS animeTitle,
      COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0) AS activeWatchedMs,
      COALESCE(asm.tokensSeen, s.tokens_seen, 0) AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount
    FROM imm_sessions s
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
    ${whereClause}
    ORDER BY s.started_at_ms ASC
  `);

  return (cutoffMs === null ? prepared.all() : prepared.all(cutoffMs)) as TrendSessionMetricRow[];
}

function buildNewWordsPerDay(db: DatabaseSync, cutoffMs: number | null): TrendChartPoint[] {
  const whereClause = cutoffMs === null ? '' : 'AND first_seen >= ?';
  const prepared = db.prepare(`
    SELECT
      CAST(first_seen / 86400 AS INTEGER) AS epochDay,
      COUNT(*) AS wordCount
    FROM imm_words
    WHERE first_seen IS NOT NULL
    ${whereClause}
    GROUP BY epochDay
    ORDER BY epochDay ASC
  `);

  const rows = (
    cutoffMs === null ? prepared.all() : prepared.all(Math.floor(cutoffMs / 1000))
  ) as Array<{
    epochDay: number;
    wordCount: number;
  }>;

  return rows.map((row) => ({
    label: dayLabel(row.epochDay),
    value: row.wordCount,
  }));
}

export function getTrendsDashboard(
  db: DatabaseSync,
  range: TrendRange = '30d',
  groupBy: TrendGroupBy = 'day',
): TrendsDashboardQueryResult {
  const dayLimit = getTrendDayLimit(range);
  const monthlyLimit = getTrendMonthlyLimit(range);
  const cutoffMs = getTrendCutoffMs(range);

  const chartRollups =
    groupBy === 'month' ? getMonthlyRollups(db, monthlyLimit) : getDailyRollups(db, dayLimit);
  const dailyRollups = getDailyRollups(db, dayLimit);
  const sessions = getTrendSessionMetrics(db, cutoffMs);
  const titlesByVideoId = getVideoAnimeTitleMap(
    db,
    dailyRollups.map((rollup) => rollup.videoId),
  );

  const aggregatedRows = buildAggregatedTrendRows(chartRollups);
  const activity = {
    watchTime: aggregatedRows.map((row) => ({ label: row.label, value: row.activeMin })),
    cards: aggregatedRows.map((row) => ({ label: row.label, value: row.cards })),
    words: aggregatedRows.map((row) => ({ label: row.label, value: row.words })),
    sessions: aggregatedRows.map((row) => ({ label: row.label, value: row.sessions })),
  };

  const animePerDay = {
    episodes: buildEpisodesPerAnimeFromDailyRollups(dailyRollups, titlesByVideoId),
    watchTime: buildPerAnimeFromDailyRollups(dailyRollups, titlesByVideoId, (rollup) =>
      Math.round(rollup.totalActiveMin),
    ),
    cards: buildPerAnimeFromDailyRollups(
      dailyRollups,
      titlesByVideoId,
      (rollup) => rollup.totalCards,
    ),
    words: buildPerAnimeFromDailyRollups(
      dailyRollups,
      titlesByVideoId,
      (rollup) => rollup.totalTokensSeen,
    ),
    lookups: buildPerAnimeFromSessions(sessions, (session) => session.yomitanLookupCount),
    lookupsPerHundred: buildLookupsPerHundredPerAnime(sessions),
  };

  return {
    activity,
    progress: {
      watchTime: accumulatePoints(activity.watchTime),
      sessions: accumulatePoints(activity.sessions),
      words: accumulatePoints(activity.words),
      newWords: accumulatePoints(buildNewWordsPerDay(db, cutoffMs)),
      cards: accumulatePoints(activity.cards),
      episodes: accumulatePoints(buildEpisodesPerDayFromDailyRollups(dailyRollups)),
      lookups: accumulatePoints(
        buildSessionSeriesByDay(sessions, (session) => session.yomitanLookupCount),
      ),
    },
    ratios: {
      lookupsPerHundred: buildLookupsPerHundredWords(sessions),
    },
    animePerDay,
    animeCumulative: {
      watchTime: buildCumulativePerAnime(animePerDay.watchTime),
      episodes: buildCumulativePerAnime(animePerDay.episodes),
      cards: buildCumulativePerAnime(animePerDay.cards),
      words: buildCumulativePerAnime(animePerDay.words),
    },
    patterns: {
      watchTimeByDayOfWeek: buildWatchTimeByDayOfWeek(sessions),
      watchTimeByHour: buildWatchTimeByHour(sessions),
    },
  };
}

export function getVocabularyStats(
  db: DatabaseSync,
  limit = 100,
  excludePos?: string[],
): VocabularyStatsRow[] {
  const hasExclude = excludePos && excludePos.length > 0;
  const placeholders = hasExclude ? excludePos.map(() => '?').join(', ') : '';
  const whereClause = hasExclude
    ? `WHERE (part_of_speech IS NULL OR part_of_speech NOT IN (${placeholders}))`
    : '';
  const stmt = db.prepare(`
    SELECT w.id AS wordId, w.headword, w.word, w.reading,
      w.part_of_speech AS partOfSpeech, w.pos1, w.pos2, w.pos3,
      w.frequency, w.frequency_rank AS frequencyRank,
      w.first_seen AS firstSeen, w.last_seen AS lastSeen,
      COUNT(DISTINCT sl.anime_id) AS animeCount
    FROM imm_words w
    LEFT JOIN imm_word_line_occurrences o ON o.word_id = w.id
    LEFT JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id AND sl.anime_id IS NOT NULL
    ${whereClause ? whereClause.replace('part_of_speech', 'w.part_of_speech') : ''}
    GROUP BY w.id
    ORDER BY w.frequency DESC LIMIT ?
  `);
  const params = hasExclude ? [...excludePos, limit] : [limit];
  return stmt.all(...params) as VocabularyStatsRow[];
}

function toStoredWordToken(row: {
  word: string;
  headword: string;
  part_of_speech: string | null;
  pos1: string | null;
  pos2: string | null;
  pos3: string | null;
}): MergedToken {
  return {
    surface: row.word || row.headword || '',
    reading: '',
    headword: row.headword || row.word || '',
    startPos: 0,
    endPos: 0,
    partOfSpeech: deriveStoredPartOfSpeech({
      partOfSpeech: row.part_of_speech,
      pos1: row.pos1,
    }),
    pos1: row.pos1 ?? '',
    pos2: row.pos2 ?? '',
    pos3: row.pos3 ?? '',
    isMerged: true,
    isKnown: false,
    isNPlusOneTarget: false,
  };
}

function normalizePosField(value: string | null | undefined): string {
  return typeof value === 'string' ? value.trim() : '';
}

function resolveStoredVocabularyPos(row: CleanupVocabularyRow): ResolvedVocabularyPos | null {
  const headword = normalizePosField(row.headword);
  const reading = normalizePosField(row.reading);
  const partOfSpeechRaw = typeof row.part_of_speech === 'string' ? row.part_of_speech.trim() : '';
  const pos1 = normalizePosField(row.pos1);
  const pos2 = normalizePosField(row.pos2);
  const pos3 = normalizePosField(row.pos3);

  if (!headword && !reading && !partOfSpeechRaw && !pos1 && !pos2 && !pos3) {
    return null;
  }

  return {
    headword: headword || normalizePosField(row.word),
    reading,
    hasPosMetadata: Boolean(partOfSpeechRaw || pos1 || pos2 || pos3),
    partOfSpeech: deriveStoredPartOfSpeech({
      partOfSpeech: partOfSpeechRaw,
      pos1,
    }),
    pos1,
    pos2,
    pos3,
  };
}

function hasStructuredPos(pos: ResolvedVocabularyPos | null): boolean {
  return Boolean(pos?.hasPosMetadata && (pos.pos1 || pos.pos2 || pos.pos3 || pos.partOfSpeech));
}

function needsLegacyVocabularyMetadataRepair(
  row: CleanupVocabularyRow,
  stored: ResolvedVocabularyPos | null,
): boolean {
  if (!stored) {
    return true;
  }

  if (!hasStructuredPos(stored)) {
    return true;
  }

  if (!stored.reading) {
    return true;
  }

  if (!stored.headword) {
    return true;
  }

  return stored.headword === normalizePosField(row.word);
}

function shouldUpdateStoredVocabularyPos(
  row: CleanupVocabularyRow,
  next: ResolvedVocabularyPos,
): boolean {
  return (
    normalizePosField(row.headword) !== next.headword ||
    normalizePosField(row.reading) !== next.reading ||
    (next.hasPosMetadata &&
      (normalizePosField(row.part_of_speech) !== next.partOfSpeech ||
        normalizePosField(row.pos1) !== next.pos1 ||
        normalizePosField(row.pos2) !== next.pos2 ||
        normalizePosField(row.pos3) !== next.pos3))
  );
}

function chooseMergedPartOfSpeech(
  current: string | null | undefined,
  incoming: ResolvedVocabularyPos,
): string {
  const normalizedCurrent = normalizePosField(current);
  if (
    normalizedCurrent &&
    normalizedCurrent !== PartOfSpeech.other &&
    incoming.partOfSpeech === PartOfSpeech.other
  ) {
    return normalizedCurrent;
  }
  return incoming.partOfSpeech;
}

async function maybeResolveLegacyVocabularyPos(
  row: CleanupVocabularyRow,
  options: CleanupVocabularyStatsOptions,
): Promise<ResolvedVocabularyPos | null> {
  const stored = resolveStoredVocabularyPos(row);
  if (!needsLegacyVocabularyMetadataRepair(row, stored) || !options.resolveLegacyPos) {
    return stored;
  }

  const resolved = await options.resolveLegacyPos(row);
  if (resolved) {
    return {
      headword: normalizePosField(resolved.headword) || normalizePosField(row.word),
      reading: normalizePosField(resolved.reading),
      hasPosMetadata: true,
      partOfSpeech: deriveStoredPartOfSpeech({
        partOfSpeech: resolved.partOfSpeech,
        pos1: resolved.pos1,
      }),
      pos1: normalizePosField(resolved.pos1),
      pos2: normalizePosField(resolved.pos2),
      pos3: normalizePosField(resolved.pos3),
    };
  }

  return stored;
}

export async function cleanupVocabularyStats(
  db: DatabaseSync,
  options: CleanupVocabularyStatsOptions = {},
): Promise<VocabularyCleanupSummary> {
  const rows = db
    .prepare(
      `SELECT id, word, headword, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
       FROM imm_words`,
    )
    .all() as CleanupVocabularyRow[];
  const findDuplicateStmt = db.prepare(
    `SELECT id, part_of_speech, pos1, pos2, pos3, first_seen, last_seen, frequency
     FROM imm_words
     WHERE headword = ? AND word = ? AND reading = ? AND id != ?`,
  );
  const deleteStmt = db.prepare('DELETE FROM imm_words WHERE id = ?');
  const updateStmt = db.prepare(
    `UPDATE imm_words
     SET headword = ?, reading = ?, part_of_speech = ?, pos1 = ?, pos2 = ?, pos3 = ?
     WHERE id = ?`,
  );
  const mergeWordStmt = db.prepare(
    `UPDATE imm_words
     SET
       frequency = COALESCE(frequency, 0) + ?,
       part_of_speech = ?,
       pos1 = ?,
       pos2 = ?,
       pos3 = ?,
       first_seen = MIN(COALESCE(first_seen, ?), ?),
       last_seen = MAX(COALESCE(last_seen, ?), ?)
     WHERE id = ?`,
  );
  const moveOccurrencesStmt = db.prepare(
    `INSERT INTO imm_word_line_occurrences (line_id, word_id, occurrence_count)
     SELECT line_id, ?, occurrence_count
     FROM imm_word_line_occurrences
     WHERE word_id = ?
     ON CONFLICT(line_id, word_id) DO UPDATE SET
       occurrence_count = imm_word_line_occurrences.occurrence_count + excluded.occurrence_count`,
  );
  const deleteOccurrencesStmt = db.prepare(
    'DELETE FROM imm_word_line_occurrences WHERE word_id = ?',
  );
  let kept = 0;
  let deleted = 0;
  let repaired = 0;

  for (const row of rows) {
    const resolvedPos = await maybeResolveLegacyVocabularyPos(row, options);
    const shouldRepair = Boolean(resolvedPos && shouldUpdateStoredVocabularyPos(row, resolvedPos));
    if (resolvedPos && shouldRepair) {
      const duplicate = findDuplicateStmt.get(
        resolvedPos.headword,
        row.word,
        resolvedPos.reading,
        row.id,
      ) as {
        id: number;
        part_of_speech: string | null;
        pos1: string | null;
        pos2: string | null;
        pos3: string | null;
        first_seen: number | null;
        last_seen: number | null;
        frequency: number | null;
      } | null;
      if (duplicate) {
        moveOccurrencesStmt.run(duplicate.id, row.id);
        deleteOccurrencesStmt.run(row.id);
        mergeWordStmt.run(
          row.frequency ?? 0,
          chooseMergedPartOfSpeech(duplicate.part_of_speech, resolvedPos),
          normalizePosField(duplicate.pos1) || resolvedPos.pos1,
          normalizePosField(duplicate.pos2) || resolvedPos.pos2,
          normalizePosField(duplicate.pos3) || resolvedPos.pos3,
          row.first_seen ?? duplicate.first_seen ?? 0,
          row.first_seen ?? duplicate.first_seen ?? 0,
          row.last_seen ?? duplicate.last_seen ?? 0,
          row.last_seen ?? duplicate.last_seen ?? 0,
          duplicate.id,
        );
        deleteStmt.run(row.id);
        repaired += 1;
        deleted += 1;
        continue;
      }

      updateStmt.run(
        resolvedPos.headword,
        resolvedPos.reading,
        resolvedPos.partOfSpeech,
        resolvedPos.pos1,
        resolvedPos.pos2,
        resolvedPos.pos3,
        row.id,
      );
      repaired += 1;
    }

    const effectiveRow = {
      ...row,
      headword: resolvedPos?.headword ?? row.headword,
      reading: resolvedPos?.reading ?? row.reading,
      part_of_speech: resolvedPos?.hasPosMetadata ? resolvedPos.partOfSpeech : row.part_of_speech,
      pos1: resolvedPos?.pos1 ?? row.pos1,
      pos2: resolvedPos?.pos2 ?? row.pos2,
      pos3: resolvedPos?.pos3 ?? row.pos3,
    };
    const missingPos =
      !normalizePosField(effectiveRow.part_of_speech) &&
      !normalizePosField(effectiveRow.pos1) &&
      !normalizePosField(effectiveRow.pos2) &&
      !normalizePosField(effectiveRow.pos3);
    if (
      missingPos ||
      shouldExcludeTokenFromVocabularyPersistence(toStoredWordToken(effectiveRow))
    ) {
      deleteStmt.run(row.id);
      deleted += 1;
      continue;
    }
    kept += 1;
  }

  return {
    scanned: rows.length,
    kept,
    deleted,
    repaired,
  };
}

export function getKanjiStats(db: DatabaseSync, limit = 100): KanjiStatsRow[] {
  const stmt = db.prepare(`
    SELECT id AS kanjiId, kanji, frequency,
      first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_kanji ORDER BY frequency DESC LIMIT ?
  `);
  return stmt.all(limit) as KanjiStatsRow[];
}

export function getWordOccurrences(
  db: DatabaseSync,
  headword: string,
  word: string,
  reading: string,
  limit = 100,
  offset = 0,
): WordOccurrenceRow[] {
  return db
    .prepare(
      `
        SELECT
          l.anime_id AS animeId,
          a.canonical_title AS animeTitle,
          l.video_id AS videoId,
          v.canonical_title AS videoTitle,
          v.source_path AS sourcePath,
          l.secondary_text AS secondaryText,
          l.session_id AS sessionId,
          l.line_index AS lineIndex,
          l.segment_start_ms AS segmentStartMs,
          l.segment_end_ms AS segmentEndMs,
          l.text AS text,
          o.occurrence_count AS occurrenceCount
        FROM imm_word_line_occurrences o
        JOIN imm_words w ON w.id = o.word_id
        JOIN imm_subtitle_lines l ON l.line_id = o.line_id
        JOIN imm_videos v ON v.video_id = l.video_id
        LEFT JOIN imm_anime a ON a.anime_id = l.anime_id
        WHERE w.headword = ? AND w.word = ? AND w.reading = ?
        ORDER BY l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
        OFFSET ?
      `,
    )
    .all(headword, word, reading, limit, offset) as unknown as WordOccurrenceRow[];
}

export function getKanjiOccurrences(
  db: DatabaseSync,
  kanji: string,
  limit = 100,
  offset = 0,
): KanjiOccurrenceRow[] {
  return db
    .prepare(
      `
        SELECT
          l.anime_id AS animeId,
          a.canonical_title AS animeTitle,
          l.video_id AS videoId,
          v.canonical_title AS videoTitle,
          v.source_path AS sourcePath,
          l.secondary_text AS secondaryText,
          l.session_id AS sessionId,
          l.line_index AS lineIndex,
          l.segment_start_ms AS segmentStartMs,
          l.segment_end_ms AS segmentEndMs,
          l.text AS text,
          o.occurrence_count AS occurrenceCount
        FROM imm_kanji_line_occurrences o
        JOIN imm_kanji k ON k.id = o.kanji_id
        JOIN imm_subtitle_lines l ON l.line_id = o.line_id
        JOIN imm_videos v ON v.video_id = l.video_id
        LEFT JOIN imm_anime a ON a.anime_id = l.anime_id
        WHERE k.kanji = ?
        ORDER BY l.CREATED_DATE DESC, l.line_id DESC
        LIMIT ?
        OFFSET ?
      `,
    )
    .all(kanji, limit, offset) as unknown as KanjiOccurrenceRow[];
}

export function getSessionEvents(
  db: DatabaseSync,
  sessionId: number,
  limit = 500,
  eventTypes?: number[],
): SessionEventRow[] {
  if (!eventTypes || eventTypes.length === 0) {
    const stmt = db.prepare(`
      SELECT event_type AS eventType, ts_ms AS tsMs, payload_json AS payload
      FROM imm_session_events WHERE session_id = ? ORDER BY ts_ms ASC LIMIT ?
    `);
    return stmt.all(sessionId, limit) as SessionEventRow[];
  }

  const placeholders = eventTypes.map(() => '?').join(', ');
  const stmt = db.prepare(`
    SELECT event_type AS eventType, ts_ms AS tsMs, payload_json AS payload
    FROM imm_session_events
    WHERE session_id = ? AND event_type IN (${placeholders})
    ORDER BY ts_ms ASC
    LIMIT ?
  `);
  return stmt.all(sessionId, ...eventTypes, limit) as SessionEventRow[];
}

export function getAnimeLibrary(db: DatabaseSync): AnimeLibraryRow[] {
  return db
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
    .all() as unknown as AnimeLibraryRow[];
}

export function getAnimeDetail(db: DatabaseSync, animeId: number): AnimeDetailRow | null {
  return db
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
    .get(animeId) as unknown as AnimeDetailRow | null;
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
  return db
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
        SELECT s_recent.ended_media_ms
        FROM imm_sessions s_recent
        WHERE s_recent.video_id = v.video_id
          AND s_recent.ended_media_ms IS NOT NULL
        ORDER BY
          COALESCE(s_recent.ended_at_ms, s_recent.LAST_UPDATE_DATE, s_recent.started_at_ms) DESC,
          s_recent.session_id DESC
        LIMIT 1
      ) AS endedMediaMs,
      v.watched AS watched,
      COUNT(DISTINCT s.session_id) AS totalSessions,
      COALESCE(SUM(COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0)), 0) AS totalActiveMs,
      COALESCE(SUM(COALESCE(asm.cardsMined, s.cards_mined, 0)), 0) AS totalCards,
      COALESCE(SUM(COALESCE(asm.tokensSeen, s.tokens_seen, 0)), 0) AS totalTokensSeen,
      COALESCE(SUM(COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0)), 0) AS totalYomitanLookupCount,
      MAX(s.started_at_ms) AS lastWatchedMs
    FROM imm_videos v
    JOIN imm_sessions s ON s.video_id = v.video_id
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
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
    .all(animeId) as unknown as AnimeEpisodeRow[];
}

export function getMediaLibrary(db: DatabaseSync): MediaLibraryRow[] {
  return db
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
      CASE
        WHEN ma.cover_blob_hash IS NOT NULL OR ma.cover_blob IS NOT NULL THEN 1
        ELSE 0
      END AS hasCoverArt
    FROM imm_videos v
    JOIN imm_lifetime_media lm ON lm.video_id = v.video_id
    LEFT JOIN imm_media_art ma ON ma.video_id = v.video_id
    ORDER BY lm.last_watched_ms DESC
  `,
    )
    .all() as unknown as MediaLibraryRow[];
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
      COALESCE(SUM(COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0)), 0) AS totalYomitanLookupCount
    FROM imm_videos v
    JOIN imm_lifetime_media lm ON lm.video_id = v.video_id
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
  return db
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
      COALESCE(asm.tokensSeen, s.tokens_seen, 0) AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.lookupCount, s.lookup_count, 0) AS lookupCount,
      COALESCE(asm.lookupHits, s.lookup_hits, 0) AS lookupHits,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount
    FROM imm_sessions s
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    WHERE s.video_id = ?
    ORDER BY s.started_at_ms DESC
    LIMIT ?
  `,
    )
    .all(videoId, limit) as unknown as SessionSummaryQueryRow[];
}

export function getMediaDailyRollups(
  db: DatabaseSync,
  videoId: number,
  limit = 90,
): ImmersionSessionRollupRow[] {
  return db
    .prepare(
      `
    WITH recent_days AS (
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
      total_tokens_seen AS totalTokensSeen,
      total_cards AS totalCards,
      cards_per_hour AS cardsPerHour,
      tokens_per_min AS tokensPerMin,
      lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups
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
  return db
    .prepare(
      `
    WITH recent_days AS (
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
           r.total_tokens_seen AS totalTokensSeen, r.total_cards AS totalCards,
           r.cards_per_hour AS cardsPerHour, r.tokens_per_min AS tokensPerMin,
           r.lookup_hit_rate AS lookupHitRate
    FROM imm_daily_rollups r
    JOIN imm_videos v ON v.video_id = r.video_id
    WHERE v.anime_id = ?
    AND r.rollup_day IN (SELECT rollup_day FROM recent_days)
    ORDER BY r.rollup_day DESC, r.video_id DESC
    `,
    )
    .all(animeId, limit, animeId) as unknown as ImmersionSessionRollupRow[];
}

export function getAnimeCoverArt(db: DatabaseSync, animeId: number): MediaArtRow | null {
  const resolvedCoverBlob = resolvedCoverBlobExpr('a', 'cab');
  return db
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
    .get(animeId) as unknown as MediaArtRow | null;
}

export function getCoverArt(db: DatabaseSync, videoId: number): MediaArtRow | null {
  const resolvedCoverBlob = resolvedCoverBlobExpr('a', 'cab');
  return db
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
    .get(videoId) as unknown as MediaArtRow | null;
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
    WHERE sl.anime_id = ?
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

export function getWordDetail(db: DatabaseSync, wordId: number): WordDetailRow | null {
  return db
    .prepare(
      `
    SELECT id AS wordId, headword, word, reading,
           part_of_speech AS partOfSpeech, pos1, pos2, pos3,
           frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_words WHERE id = ?
  `,
    )
    .get(wordId) as WordDetailRow | null;
}

export function getWordAnimeAppearances(
  db: DatabaseSync,
  wordId: number,
): WordAnimeAppearanceRow[] {
  return db
    .prepare(
      `
    SELECT a.anime_id AS animeId, a.canonical_title AS animeTitle,
           SUM(o.occurrence_count) AS occurrenceCount
    FROM imm_word_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE o.word_id = ? AND sl.anime_id IS NOT NULL
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `,
    )
    .all(wordId) as WordAnimeAppearanceRow[];
}

export function getSimilarWords(db: DatabaseSync, wordId: number, limit = 10): SimilarWordRow[] {
  const word = db.prepare('SELECT headword, reading FROM imm_words WHERE id = ?').get(wordId) as {
    headword: string;
    reading: string;
  } | null;
  if (!word) return [];
  return db
    .prepare(
      `
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE id != ?
    AND (reading = ? OR headword LIKE ? OR headword LIKE ?)
    ORDER BY frequency DESC
    LIMIT ?
  `,
    )
    .all(
      wordId,
      word.reading,
      `%${word.headword.charAt(0)}%`,
      `%${word.headword.charAt(word.headword.length - 1)}%`,
      limit,
    ) as SimilarWordRow[];
}

export function getKanjiDetail(db: DatabaseSync, kanjiId: number): KanjiDetailRow | null {
  return db
    .prepare(
      `
    SELECT id AS kanjiId, kanji, frequency, first_seen AS firstSeen, last_seen AS lastSeen
    FROM imm_kanji WHERE id = ?
  `,
    )
    .get(kanjiId) as KanjiDetailRow | null;
}

export function getKanjiAnimeAppearances(
  db: DatabaseSync,
  kanjiId: number,
): KanjiAnimeAppearanceRow[] {
  return db
    .prepare(
      `
    SELECT a.anime_id AS animeId, a.canonical_title AS animeTitle,
           SUM(o.occurrence_count) AS occurrenceCount
    FROM imm_kanji_line_occurrences o
    JOIN imm_subtitle_lines sl ON sl.line_id = o.line_id
    JOIN imm_anime a ON a.anime_id = sl.anime_id
    WHERE o.kanji_id = ? AND sl.anime_id IS NOT NULL
    GROUP BY a.anime_id
    ORDER BY occurrenceCount DESC
  `,
    )
    .all(kanjiId) as KanjiAnimeAppearanceRow[];
}

export function getKanjiWords(db: DatabaseSync, kanjiId: number, limit = 20): KanjiWordRow[] {
  const kanjiRow = db.prepare('SELECT kanji FROM imm_kanji WHERE id = ?').get(kanjiId) as {
    kanji: string;
  } | null;
  if (!kanjiRow) return [];
  return db
    .prepare(
      `
    SELECT id AS wordId, headword, word, reading, frequency
    FROM imm_words
    WHERE headword LIKE ?
    ORDER BY frequency DESC
    LIMIT ?
  `,
    )
    .all(`%${kanjiRow.kanji}%`, limit) as KanjiWordRow[];
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
  return db
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
      COALESCE(asm.tokensSeen, s.tokens_seen, 0) AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.lookupCount, s.lookup_count, 0) AS lookupCount,
      COALESCE(asm.lookupHits, s.lookup_hits, 0) AS lookupHits,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount
    FROM imm_sessions s
    JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    WHERE s.video_id = ?
    ORDER BY s.started_at_ms DESC
  `,
    )
    .all(videoId) as SessionSummaryQueryRow[];
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
    ORDER BY e.ts_ms DESC
  `,
    )
    .all(videoId) as Array<{
    eventId: number;
    sessionId: number;
    tsMs: number;
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
      tsMs: row.tsMs,
      cardsDelta: row.cardsDelta,
      noteIds,
    };
  });
}

export function upsertCoverArt(
  db: DatabaseSync,
  videoId: number,
  art: {
    anilistId: number | null;
    coverUrl: string | null;
    coverBlob: ArrayBuffer | Uint8Array | Buffer | null;
    titleRomaji: string | null;
    titleEnglish: string | null;
    episodesTotal: number | null;
  },
): void {
  const existing = db
    .prepare(
      `
        SELECT cover_blob_hash AS coverBlobHash
        FROM imm_media_art
        WHERE video_id = ?
      `,
    )
    .get(videoId) as { coverBlobHash: string | null } | undefined;
  const sharedCoverBlobHash = findSharedCoverBlobHash(db, videoId, art.anilistId, art.coverUrl);
  const nowMs = Date.now();
  const coverBlob = normalizeCoverBlobBytes(art.coverBlob);
  let coverBlobHash = sharedCoverBlobHash ?? null;
  if (!coverBlobHash && coverBlob && coverBlob.length > 0) {
    coverBlobHash = createHash('sha256').update(coverBlob).digest('hex');
  }
  if (!coverBlobHash && (!coverBlob || coverBlob.length === 0)) {
    coverBlobHash = existing?.coverBlobHash ?? null;
  }

  if (coverBlobHash && coverBlob && coverBlob.length > 0 && !sharedCoverBlobHash) {
    db.prepare(
      `
        INSERT INTO imm_cover_art_blobs (blob_hash, cover_blob, CREATED_DATE, LAST_UPDATE_DATE)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(blob_hash) DO UPDATE SET
          LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
      `,
    ).run(coverBlobHash, coverBlob, nowMs, nowMs);
  }

  db.prepare(
    `
    INSERT INTO imm_media_art (
      video_id, anilist_id, cover_url, cover_blob, cover_blob_hash,
      title_romaji, title_english, episodes_total,
      fetched_at_ms, CREATED_DATE, LAST_UPDATE_DATE
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(video_id) DO UPDATE SET
      anilist_id = excluded.anilist_id,
      cover_url = excluded.cover_url,
      cover_blob = excluded.cover_blob,
      cover_blob_hash = excluded.cover_blob_hash,
      title_romaji = excluded.title_romaji,
      title_english = excluded.title_english,
      episodes_total = excluded.episodes_total,
      fetched_at_ms = excluded.fetched_at_ms,
      LAST_UPDATE_DATE = excluded.LAST_UPDATE_DATE
  `,
  ).run(
    videoId,
    art.anilistId,
    art.coverUrl,
    coverBlobHash ? buildCoverBlobReference(coverBlobHash) : coverBlob,
    coverBlobHash,
    art.titleRomaji,
    art.titleEnglish,
    art.episodesTotal,
    nowMs,
    nowMs,
    nowMs,
  );

  if (existing?.coverBlobHash !== coverBlobHash) {
    cleanupUnusedCoverArtBlobHash(db, existing?.coverBlobHash ?? null);
  }
}

export function updateAnimeAnilistInfo(
  db: DatabaseSync,
  videoId: number,
  info: {
    anilistId: number;
    titleRomaji: string | null;
    titleEnglish: string | null;
    titleNative: string | null;
    episodesTotal: number | null;
  },
): void {
  const row = db.prepare('SELECT anime_id FROM imm_videos WHERE video_id = ?').get(videoId) as {
    anime_id: number | null;
  } | null;
  if (!row?.anime_id) return;

  db.prepare(
    `
    UPDATE imm_anime
    SET
      anilist_id = COALESCE(?, anilist_id),
      title_romaji = COALESCE(?, title_romaji),
      title_english = COALESCE(?, title_english),
      title_native = COALESCE(?, title_native),
      episodes_total = COALESCE(?, episodes_total),
      LAST_UPDATE_DATE = ?
    WHERE anime_id = ?
  `,
  ).run(
    info.anilistId,
    info.titleRomaji,
    info.titleEnglish,
    info.titleNative,
    info.episodesTotal,
    Date.now(),
    row.anime_id,
  );
}

export function markVideoWatched(db: DatabaseSync, videoId: number, watched: boolean): void {
  db.prepare('UPDATE imm_videos SET watched = ?, LAST_UPDATE_DATE = ? WHERE video_id = ?').run(
    watched ? 1 : 0,
    Date.now(),
    videoId,
  );
}

export function getVideoDurationMs(db: DatabaseSync, videoId: number): number {
  const row = db.prepare('SELECT duration_ms FROM imm_videos WHERE video_id = ?').get(videoId) as {
    duration_ms: number;
  } | null;
  return row?.duration_ms ?? 0;
}

export function isVideoWatched(db: DatabaseSync, videoId: number): boolean {
  const row = db.prepare('SELECT watched FROM imm_videos WHERE video_id = ?').get(videoId) as {
    watched: number;
  } | null;
  return row?.watched === 1;
}

export function deleteSession(db: DatabaseSync, sessionId: number): void {
  const sessionIds = [sessionId];
  const affectedWordIds = getAffectedWordIdsForSessions(db, sessionIds);
  const affectedKanjiIds = getAffectedKanjiIdsForSessions(db, sessionIds);

  db.exec('BEGIN IMMEDIATE');
  try {
    deleteSessionsByIds(db, sessionIds);
    refreshLexicalAggregates(db, affectedWordIds, affectedKanjiIds);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

export function deleteSessions(db: DatabaseSync, sessionIds: number[]): void {
  if (sessionIds.length === 0) return;
  const affectedWordIds = getAffectedWordIdsForSessions(db, sessionIds);
  const affectedKanjiIds = getAffectedKanjiIdsForSessions(db, sessionIds);

  db.exec('BEGIN IMMEDIATE');
  try {
    deleteSessionsByIds(db, sessionIds);
    refreshLexicalAggregates(db, affectedWordIds, affectedKanjiIds);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}

export function deleteVideo(db: DatabaseSync, videoId: number): void {
  const artRow = db
    .prepare(
      `
        SELECT cover_blob_hash AS coverBlobHash
        FROM imm_media_art
        WHERE video_id = ?
      `,
    )
    .get(videoId) as { coverBlobHash: string | null } | undefined;
  const affectedWordIds = getAffectedWordIdsForVideo(db, videoId);
  const affectedKanjiIds = getAffectedKanjiIdsForVideo(db, videoId);
  const sessions = db
    .prepare('SELECT session_id FROM imm_sessions WHERE video_id = ?')
    .all(videoId) as Array<{ session_id: number }>;

  db.exec('BEGIN IMMEDIATE');
  try {
    deleteSessionsByIds(
      db,
      sessions.map((session) => session.session_id),
    );
    db.prepare('DELETE FROM imm_subtitle_lines WHERE video_id = ?').run(videoId);
    db.prepare('DELETE FROM imm_daily_rollups WHERE video_id = ?').run(videoId);
    db.prepare('DELETE FROM imm_monthly_rollups WHERE video_id = ?').run(videoId);
    db.prepare('DELETE FROM imm_media_art WHERE video_id = ?').run(videoId);
    cleanupUnusedCoverArtBlobHash(db, artRow?.coverBlobHash ?? null);
    db.prepare('DELETE FROM imm_videos WHERE video_id = ?').run(videoId);
    refreshLexicalAggregates(db, affectedWordIds, affectedKanjiIds);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
}
