import type { DatabaseSync } from './sqlite';
import type { ImmersionSessionRollupRow } from './types';
import {
  ACTIVE_SESSION_METRICS_CTE,
  currentDbTimestamp,
  getLocalDayOfWeek,
  getLocalEpochDay,
  getLocalHourOfDay,
  getLocalMonthKey,
  getShiftedLocalDayTimestamp,
  makePlaceholders,
  sessionDisplayWordsExpr,
  toDbTimestamp,
} from './query-shared';
import { getDailyRollups, getMonthlyRollups } from './query-sessions';

type TrendRange = '7d' | '30d' | '90d' | '365d' | 'all';
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

export interface LibrarySummaryRow {
  title: string;
  watchTimeMin: number;
  videos: number;
  sessions: number;
  cards: number;
  words: number;
  lookups: number;
  lookupsPerHundred: number | null;
  firstWatched: number;
  lastWatched: number;
}

interface TrendSessionMetricRow {
  startedAtMs: number;
  epochDay: number;
  monthKey: number;
  dayOfWeek: number;
  hourOfDay: number;
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
    cardsPerHour: TrendChartPoint[];
    readingSpeed: TrendChartPoint[];
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
  librarySummary: LibrarySummaryRow[];
}

const TREND_DAY_LIMITS: Record<Exclude<TrendRange, 'all'>, number> = {
  '7d': 7,
  '30d': 30,
  '90d': 90,
  '365d': 365,
};

const MONTH_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

function getTrendDayLimit(range: TrendRange): number {
  return range === 'all' ? 365 : TREND_DAY_LIMITS[range];
}

function getTrendMonthlyLimit(db: DatabaseSync, range: TrendRange): number {
  if (range === 'all') {
    return 120;
  }
  const currentTimestamp = currentDbTimestamp();
  const todayStartMs = getShiftedLocalDayTimestamp(db, currentTimestamp, 0);
  const cutoffMs = getShiftedLocalDayTimestamp(
    db,
    currentTimestamp,
    -(TREND_DAY_LIMITS[range] - 1),
  );
  const currentMonthKey = getLocalMonthKey(db, todayStartMs);
  const cutoffMonthKey = getLocalMonthKey(db, cutoffMs);
  const currentYear = Math.floor(currentMonthKey / 100);
  const currentMonth = currentMonthKey % 100;
  const cutoffYear = Math.floor(cutoffMonthKey / 100);
  const cutoffMonth = cutoffMonthKey % 100;
  return Math.max(1, (currentYear - cutoffYear) * 12 + currentMonth - cutoffMonth + 1);
}

function getTrendCutoffMs(db: DatabaseSync, range: TrendRange): string | null {
  if (range === 'all') {
    return null;
  }
  return getShiftedLocalDayTimestamp(db, currentDbTimestamp(), -(getTrendDayLimit(range) - 1));
}

function dayPartsFromEpochDay(epochDay: number): { year: number; month: number; day: number } {
  const z = epochDay + 719468;
  const era = Math.floor(z / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor(
    (doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365,
  );
  let year = yoe + era * 400;
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const day = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const month = mp < 10 ? mp + 3 : mp - 9;
  if (month <= 2) {
    year += 1;
  }
  return { year, month, day };
}

function makeTrendLabel(value: number): string {
  if (value > 100_000) {
    const year = Math.floor(value / 100);
    const month = value % 100;
    return `${MONTH_NAMES[month - 1]} ${String(year).slice(-2)}`;
  }

  const { month, day } = dayPartsFromEpochDay(value);
  return `${MONTH_NAMES[month - 1]} ${day}`;
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
    existing.activeMin += rollup.totalActiveMin;
    existing.cards += rollup.totalCards;
    existing.words += rollup.totalTokensSeen;
    existing.sessions += rollup.totalSessions;
    byKey.set(rollup.rollupDayOrMonth, existing);
  }

  return Array.from(byKey.entries())
    .sort(([left], [right]) => left - right)
    .map(([key, value]) => ({
      label: makeTrendLabel(key),
      activeMin: Math.round(value.activeMin),
      cards: value.cards,
      words: value.words,
      sessions: value.sessions,
    }));
}

function buildEfficiencyRates(rows: ReturnType<typeof buildAggregatedTrendRows>): {
  cardsPerHour: TrendChartPoint[];
  readingSpeed: TrendChartPoint[];
} {
  const cardsPerHour: TrendChartPoint[] = [];
  const readingSpeed: TrendChartPoint[] = [];
  for (const row of rows) {
    const hours = row.activeMin / 60;
    cardsPerHour.push({
      label: row.label,
      value: hours > 0 ? +(row.cards / hours).toFixed(1) : 0,
    });
    readingSpeed.push({
      label: row.label,
      value: row.activeMin > 0 ? +(row.words / row.activeMin).toFixed(1) : 0,
    });
  }
  return { cardsPerHour, readingSpeed };
}

function buildWatchTimeByDayOfWeek(sessions: TrendSessionMetricRow[]): TrendChartPoint[] {
  const totals = new Array(7).fill(0);
  for (const session of sessions) {
    totals[session.dayOfWeek] += session.activeWatchedMs;
  }
  return DAY_NAMES.map((name, index) => ({
    label: name,
    value: Math.round(totals[index] / 60_000),
  }));
}

function buildWatchTimeByHour(sessions: TrendSessionMetricRow[]): TrendChartPoint[] {
  const totals = new Array(24).fill(0);
  for (const session of sessions) {
    totals[session.hourOfDay] += session.activeWatchedMs;
  }
  return totals.map((ms, index) => ({
    label: `${String(index).padStart(2, '0')}:00`,
    value: Math.round(ms / 60_000),
  }));
}

function dayLabel(epochDay: number): string {
  const { month, day } = dayPartsFromEpochDay(epochDay);
  return `${MONTH_NAMES[month - 1]} ${day}`;
}

function buildSessionSeriesByDay(
  sessions: TrendSessionMetricRow[],
  getValue: (session: TrendSessionMetricRow) => number,
): TrendChartPoint[] {
  const byDay = new Map<number, number>();
  for (const session of sessions) {
    byDay.set(session.epochDay, (byDay.get(session.epochDay) ?? 0) + getValue(session));
  }
  return Array.from(byDay.entries())
    .sort(([left], [right]) => left - right)
    .map(([epochDay, value]) => ({ label: dayLabel(epochDay), value }));
}

function buildSessionSeriesByMonth(
  sessions: TrendSessionMetricRow[],
  getValue: (session: TrendSessionMetricRow) => number,
): TrendChartPoint[] {
  const byMonth = new Map<number, number>();
  for (const session of sessions) {
    byMonth.set(session.monthKey, (byMonth.get(session.monthKey) ?? 0) + getValue(session));
  }
  return Array.from(byMonth.entries())
    .sort(([left], [right]) => left - right)
    .map(([monthKey, value]) => ({ label: makeTrendLabel(monthKey), value }));
}

function buildLookupsPerHundredWords(
  sessions: TrendSessionMetricRow[],
  groupBy: TrendGroupBy,
): TrendChartPoint[] {
  const lookupsByBucket = new Map<number, number>();
  const wordsByBucket = new Map<number, number>();

  for (const session of sessions) {
    const bucketKey = groupBy === 'month' ? session.monthKey : session.epochDay;
    lookupsByBucket.set(
      bucketKey,
      (lookupsByBucket.get(bucketKey) ?? 0) + session.yomitanLookupCount,
    );
    wordsByBucket.set(
      bucketKey,
      (wordsByBucket.get(bucketKey) ?? 0) + getTrendSessionWordCount(session),
    );
  }

  return Array.from(lookupsByBucket.entries())
    .sort(([left], [right]) => left - right)
    .map(([bucketKey, lookups]) => {
      const words = wordsByBucket.get(bucketKey) ?? 0;
      return {
        label: groupBy === 'month' ? makeTrendLabel(bucketKey) : dayLabel(bucketKey),
        value: words > 0 ? +((lookups / words) * 100).toFixed(1) : 0,
      };
    });
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

function buildLibrarySummary(
  rollups: ImmersionSessionRollupRow[],
  sessions: TrendSessionMetricRow[],
  titlesByVideoId: Map<number, string>,
): LibrarySummaryRow[] {
  type Accum = {
    watchTimeMin: number;
    videos: Set<number>;
    cards: number;
    words: number;
    firstWatched: number;
    lastWatched: number;
    sessions: number;
    lookups: number;
  };

  const byTitle = new Map<string, Accum>();

  const ensure = (title: string): Accum => {
    const existing = byTitle.get(title);
    if (existing) return existing;
    const created: Accum = {
      watchTimeMin: 0,
      videos: new Set<number>(),
      cards: 0,
      words: 0,
      firstWatched: Number.POSITIVE_INFINITY,
      lastWatched: Number.NEGATIVE_INFINITY,
      sessions: 0,
      lookups: 0,
    };
    byTitle.set(title, created);
    return created;
  };

  for (const rollup of rollups) {
    if (rollup.videoId === null) continue;
    const title = resolveVideoAnimeTitle(rollup.videoId, titlesByVideoId);
    const acc = ensure(title);
    acc.watchTimeMin += rollup.totalActiveMin;
    acc.cards += rollup.totalCards;
    acc.words += rollup.totalTokensSeen;
    acc.videos.add(rollup.videoId);
    if (rollup.rollupDayOrMonth < acc.firstWatched) {
      acc.firstWatched = rollup.rollupDayOrMonth;
    }
    if (rollup.rollupDayOrMonth > acc.lastWatched) {
      acc.lastWatched = rollup.rollupDayOrMonth;
    }
  }

  for (const session of sessions) {
    const title = resolveTrendAnimeTitle(session);
    if (!byTitle.has(title)) continue;
    const acc = byTitle.get(title)!;
    acc.sessions += 1;
    acc.lookups += session.yomitanLookupCount;
  }

  const rows: LibrarySummaryRow[] = [];
  for (const [title, acc] of byTitle) {
    if (!Number.isFinite(acc.firstWatched) || !Number.isFinite(acc.lastWatched)) {
      continue;
    }
    rows.push({
      title,
      watchTimeMin: Math.round(acc.watchTimeMin),
      videos: acc.videos.size,
      sessions: acc.sessions,
      cards: acc.cards,
      words: acc.words,
      lookups: acc.lookups,
      lookupsPerHundred: acc.words > 0 ? +((acc.lookups / acc.words) * 100).toFixed(1) : null,
      firstWatched: acc.firstWatched,
      lastWatched: acc.lastWatched,
    });
  }

  rows.sort((a, b) => b.watchTimeMin - a.watchTimeMin || a.title.localeCompare(b.title));
  return rows;
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

function buildEpisodesPerMonthFromRollups(rollups: ImmersionSessionRollupRow[]): TrendChartPoint[] {
  const byMonth = new Map<number, Set<number>>();

  for (const rollup of rollups) {
    if (rollup.videoId === null) {
      continue;
    }
    const videoIds = byMonth.get(rollup.rollupDayOrMonth) ?? new Set<number>();
    videoIds.add(rollup.videoId);
    byMonth.set(rollup.rollupDayOrMonth, videoIds);
  }

  return Array.from(byMonth.entries())
    .sort(([left], [right]) => left - right)
    .map(([monthKey, videoIds]) => ({
      label: makeTrendLabel(monthKey),
      value: videoIds.size,
    }));
}

function getTrendSessionMetrics(
  db: DatabaseSync,
  cutoffMs: string | null,
): TrendSessionMetricRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'swc', 'COALESCE(asm.tokensSeen, s.tokens_seen)');
  const whereClause = cutoffMs === null ? '' : 'WHERE s.started_at_ms >= ?';
  const cutoffValue = cutoffMs === null ? null : toDbTimestamp(cutoffMs);
  const prepared = db.prepare(`
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      s.started_at_ms AS startedAtMs,
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      a.canonical_title AS animeTitle,
      COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0) AS activeWatchedMs,
      ${wordsExpr} AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount
    FROM imm_sessions s
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN session_word_counts swc ON swc.sessionId = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
    ${whereClause}
    ORDER BY s.started_at_ms ASC
  `);

  const rows = (cutoffValue === null ? prepared.all() : prepared.all(cutoffValue)) as Array<
    TrendSessionMetricRow & { startedAtMs: number | string }
  >;
  return rows.map((row) => ({
    ...row,
    startedAtMs: 0,
    epochDay: getLocalEpochDay(db, row.startedAtMs),
    monthKey: getLocalMonthKey(db, row.startedAtMs),
    dayOfWeek: getLocalDayOfWeek(db, row.startedAtMs),
    hourOfDay: getLocalHourOfDay(db, row.startedAtMs),
  }));
}

function buildNewWordsPerDay(db: DatabaseSync, cutoffMs: string | null): TrendChartPoint[] {
  const whereClause = cutoffMs === null ? '' : 'AND first_seen >= ?';
  const prepared = db.prepare(`
    SELECT
      CAST(
        julianday(CAST(first_seen AS REAL), 'unixepoch', 'localtime') - 2440587.5
        AS INTEGER
      ) AS epochDay,
      COUNT(*) AS wordCount
    FROM imm_words
    WHERE first_seen IS NOT NULL
    ${whereClause}
    GROUP BY epochDay
    ORDER BY epochDay ASC
  `);

  const rows = (
    cutoffMs === null ? prepared.all() : prepared.all((BigInt(cutoffMs) / 1000n).toString())
  ) as Array<{
    epochDay: number;
    wordCount: number;
  }>;

  return rows.map((row) => ({
    label: dayLabel(row.epochDay),
    value: row.wordCount,
  }));
}

function buildNewWordsPerMonth(db: DatabaseSync, cutoffMs: string | null): TrendChartPoint[] {
  const whereClause = cutoffMs === null ? '' : 'AND first_seen >= ?';
  const prepared = db.prepare(`
    SELECT
      CAST(
        strftime('%Y%m', CAST(first_seen AS REAL), 'unixepoch', 'localtime')
        AS INTEGER
      ) AS monthKey,
      COUNT(*) AS wordCount
    FROM imm_words
    WHERE first_seen IS NOT NULL
    ${whereClause}
    GROUP BY monthKey
    ORDER BY monthKey ASC
  `);

  const rows = (
    cutoffMs === null ? prepared.all() : prepared.all((BigInt(cutoffMs) / 1000n).toString())
  ) as Array<{
    monthKey: number;
    wordCount: number;
  }>;

  return rows.map((row) => ({
    label: makeTrendLabel(row.monthKey),
    value: row.wordCount,
  }));
}

export function getTrendsDashboard(
  db: DatabaseSync,
  range: TrendRange = '30d',
  groupBy: TrendGroupBy = 'day',
): TrendsDashboardQueryResult {
  const dayLimit = getTrendDayLimit(range);
  const monthlyLimit = getTrendMonthlyLimit(db, range);
  const cutoffMs = getTrendCutoffMs(db, range);
  const useMonthlyBuckets = groupBy === 'month';
  const dailyRollups = getDailyRollups(db, dayLimit);
  const monthlyRollups = getMonthlyRollups(db, monthlyLimit);

  const chartRollups = useMonthlyBuckets ? monthlyRollups : dailyRollups;
  const sessions = getTrendSessionMetrics(db, cutoffMs);
  const titlesByVideoId = getVideoAnimeTitleMap(
    db,
    dailyRollups.map((rollup) => rollup.videoId),
  );

  const aggregatedRows = buildAggregatedTrendRows(chartRollups);
  const efficiency = buildEfficiencyRates(aggregatedRows);
  const activity = {
    watchTime: aggregatedRows.map((row) => ({ label: row.label, value: row.activeMin })),
    cards: aggregatedRows.map((row) => ({ label: row.label, value: row.cards })),
    words: aggregatedRows.map((row) => ({ label: row.label, value: row.words })),
    sessions: aggregatedRows.map((row) => ({ label: row.label, value: row.sessions })),
  };

  const animePerDay = {
    episodes: buildEpisodesPerAnimeFromDailyRollups(dailyRollups, titlesByVideoId),
    watchTime: buildPerAnimeFromDailyRollups(
      dailyRollups,
      titlesByVideoId,
      (rollup) => rollup.totalActiveMin,
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
  };

  return {
    activity,
    progress: {
      watchTime: accumulatePoints(activity.watchTime),
      sessions: accumulatePoints(activity.sessions),
      words: accumulatePoints(activity.words),
      newWords: accumulatePoints(
        useMonthlyBuckets ? buildNewWordsPerMonth(db, cutoffMs) : buildNewWordsPerDay(db, cutoffMs),
      ),
      cards: accumulatePoints(activity.cards),
      episodes: accumulatePoints(
        useMonthlyBuckets
          ? buildEpisodesPerMonthFromRollups(monthlyRollups)
          : buildEpisodesPerDayFromDailyRollups(dailyRollups),
      ),
      lookups: accumulatePoints(
        useMonthlyBuckets
          ? buildSessionSeriesByMonth(sessions, (session) => session.yomitanLookupCount)
          : buildSessionSeriesByDay(sessions, (session) => session.yomitanLookupCount),
      ),
    },
    ratios: {
      lookupsPerHundred: buildLookupsPerHundredWords(sessions, groupBy),
      cardsPerHour: efficiency.cardsPerHour,
      readingSpeed: efficiency.readingSpeed,
    },
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
    librarySummary: buildLibrarySummary(dailyRollups, sessions, titlesByVideoId),
  };
}
