import type { DatabaseSync } from './sqlite';
import type { ImmersionSessionRollupRow } from './types';
import { ACTIVE_SESSION_METRICS_CTE, makePlaceholders } from './query-shared.js';
import { getDailyRollups, getMonthlyRollups } from './query-sessions.js';

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

function getLocalEpochDay(timestampMs: number): number {
  const date = new Date(timestampMs);
  return Math.floor((timestampMs - date.getTimezoneOffset() * 60_000) / 86_400_000);
}

function getLocalDateForEpochDay(epochDay: number): Date {
  const utcDate = new Date(epochDay * 86_400_000);
  return new Date(utcDate.getTime() + utcDate.getTimezoneOffset() * 60_000);
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
  return getLocalDateForEpochDay(epochDay).toLocaleDateString(undefined, {
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
    const epochDay = getLocalEpochDay(session.startedAtMs);
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
    const epochDay = getLocalEpochDay(session.startedAtMs);
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
    const epochDay = getLocalEpochDay(session.startedAtMs);
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
    const epochDay = getLocalEpochDay(session.startedAtMs);

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
      CAST(julianday(first_seen, 'unixepoch', 'localtime') - 2440587.5 AS INTEGER) AS epochDay,
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
