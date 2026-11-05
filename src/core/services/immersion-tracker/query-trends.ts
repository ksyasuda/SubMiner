import type {
  StatsTrendGroupBy,
  StatsTrendPresetRange,
  StatsTrendRange,
} from '../../../types/stats-http-contract';
import {
  TREND_FREQUENCY_BANDS,
  type LibrarySummaryRow,
  type TrendBucket,
  type TrendChartPoint,
  type TrendFrequencyBandPoint,
  type TrendPerAnimePoint,
  type TrendRatePoint,
  type TrendSummaryTotals,
  type TrendsDashboardData,
} from '../../../types/stats-wire';
import type { DatabaseSync } from './sqlite';
import type { ImmersionSessionRollupRow } from './types';
import {
  ACTIVE_SESSION_METRICS_CTE,
  currentDbTimestamp,
  getLocalEpochDay,
  makePlaceholders,
  sessionDisplayWordsExpr,
} from './query-shared';
import {
  addRollupWordCount,
  applyRollupWordCounts,
  selectDailyRollupRowsBetween,
  selectMonthlyRollupRowsBetween,
  sessionRollupDaySql,
  sessionRollupMonthSql,
  type RollupWordCounts,
} from './query-sessions';
import { getNewWordFirstSeenRows, type NewWordFirstSeenRow } from './query-lexical';

interface TrendSessionRow {
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
  ended: number;
}

/** The dated numbers one day or bucket adds up. */
interface TrendTotals {
  activeMin: number;
  cards: number;
  words: number;
  lookups: number;
  episodes: number;
  newWords: number;
}

interface DayRange {
  startDay: number;
  endDay: number;
}

interface TrendWindow {
  /** Null for `all`, whose bounds come from the data. */
  startDay: number | null;
  endDay: number;
  today: number;
  previous: DayRange | null;
  /** Earliest day any metric needs, covering `previous` and the rolling-rate lookback. */
  loadFromDay: number | null;
}

interface TitleRef {
  title: string;
  animeId: number | null;
}

const TREND_DAY_LIMITS: Record<Exclude<StatsTrendPresetRange, 'all'>, number> = {
  '7d': 7,
  '30d': 30,
  '90d': 90,
  '365d': 365,
};

// A single day's rate is mostly noise: one card in a two-minute day reads as 30
// cards per hour. Daily charts rate a trailing week instead, and every rate needs
// this much underlying data before it is shown at all.
const ROLLING_RATE_DAYS = 7;
const MIN_RATE_ACTIVE_MIN = 30;
const MIN_RATE_WORDS = 2_000;

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

function emptyTotals(): TrendTotals {
  return { activeMin: 0, cards: 0, words: 0, lookups: 0, episodes: 0, newWords: 0 };
}

function dayToDate(epochDay: number): Date {
  return new Date(epochDay * 86_400_000);
}

function dayToMonthKey(epochDay: number): number {
  const date = dayToDate(epochDay);
  return date.getUTCFullYear() * 100 + date.getUTCMonth() + 1;
}

function monthDayBounds(monthKey: number): DayRange {
  const year = Math.floor(monthKey / 100);
  const month = monthKey % 100;
  return {
    startDay: Date.UTC(year, month - 1, 1) / 86_400_000,
    endDay: Date.UTC(year, month, 0) / 86_400_000,
  };
}

/** Bucket key for a day: the day itself, the Monday starting its week, or its `YYYYMM`. */
function bucketKeyForDay(epochDay: number, groupBy: StatsTrendGroupBy): number {
  if (groupBy === 'day') return epochDay;
  if (groupBy === 'month') return dayToMonthKey(epochDay);
  // Epoch day 0 was a Thursday, so +3 makes Monday the zero of the week.
  return epochDay - ((((epochDay + 3) % 7) + 7) % 7);
}

function formatBucketLabel(
  startDay: number,
  groupBy: StatsTrendGroupBy,
  showYear: boolean,
): string {
  const date = dayToDate(startDay);
  const month = MONTH_NAMES[date.getUTCMonth()];
  if (groupBy === 'month') return `${month} ${date.getUTCFullYear()}`;
  const label = `${month} ${date.getUTCDate()}`;
  return showYear ? `${label}, ${date.getUTCFullYear()}` : label;
}

function resolveTrendWindow(db: DatabaseSync, range: StatsTrendRange): TrendWindow {
  const today = getLocalEpochDay(db, currentDbTimestamp());
  if (range === 'all') {
    return { startDay: null, endDay: today, today, previous: null, loadFromDay: null };
  }
  const startDay =
    typeof range === 'string' ? today - (TREND_DAY_LIMITS[range] - 1) : range.startDay;
  const endDay = typeof range === 'string' ? today : Math.min(range.endDay, today);
  const span = Math.max(1, endDay - startDay + 1);
  const previous = { startDay: startDay - span, endDay: startDay - 1 };
  return {
    startDay,
    endDay,
    today,
    previous,
    loadFromDay: Math.min(previous.startDay, startDay - (ROLLING_RATE_DAYS - 1)),
  };
}

function buildBuckets(
  firstDay: number,
  lastDay: number,
  groupBy: StatsTrendGroupBy,
  keepKeys: ReadonlySet<number> | null,
): Array<TrendBucket & { key: number }> {
  const showYear = dayToDate(firstDay).getUTCFullYear() !== dayToDate(lastDay).getUTCFullYear();
  const buckets = new Map<number, TrendBucket & { key: number }>();
  for (let day = firstDay; day <= lastDay; day += 1) {
    const key = bucketKeyForDay(day, groupBy);
    const existing = buckets.get(key);
    if (existing) {
      existing.endDay = day;
    } else {
      buckets.set(key, {
        key,
        label: formatBucketLabel(day, groupBy, showYear),
        startDay: day,
        endDay: day,
      });
    }
  }
  const all = [...buckets.values()];
  return keepKeys ? all.filter((bucket) => keepKeys.has(bucket.key)) : all;
}

function looksLikeJellyfinStreamTitle(title: string): boolean {
  const lowered = title.toLowerCase();
  const hasApiKey = /api[\s_-]*key(?:\s|=|$)/i.test(title);
  return (
    hasApiKey &&
    (lowered.includes('stream?') ||
      lowered.includes('/stream?') ||
      lowered.includes('/videos/') ||
      lowered.includes('mediasourceid'))
  );
}

function sanitizeTrendTitle(title: string): string {
  const normalized = title.trim();
  if (!normalized) {
    return 'Unknown';
  }
  return looksLikeJellyfinStreamTitle(normalized) ? 'Jellyfin Video' : normalized;
}

function resolveSessionTitle(session: TrendSessionRow): string {
  return sanitizeTrendTitle(session.animeTitle ?? session.canonicalTitle ?? 'Unknown');
}

function getVideoTitleMap(db: DatabaseSync, videoIds: Array<number | null>): Map<number, TitleRef> {
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
          v.anime_id AS animeId,
          COALESCE(a.canonical_title, v.canonical_title, 'Unknown') AS animeTitle
        FROM imm_videos v
        LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
        WHERE v.video_id IN (${makePlaceholders(uniqueIds)})
      `,
    )
    .all(...uniqueIds) as Array<{ videoId: number; animeId: number | null; animeTitle: string }>;

  return new Map(
    rows.map((row) => [
      row.videoId,
      { title: sanitizeTrendTitle(row.animeTitle), animeId: row.animeId },
    ]),
  );
}

/**
 * One pass over the sessions the dashboard needs. Each row's display word count
 * also feeds the rollup word merge, so it is computed once per request.
 */
function getTrendSessionRows(
  db: DatabaseSync,
  startDay: number | null,
  endDay: number,
): TrendSessionRow[] {
  const wordsExpr = sessionDisplayWordsExpr('s', 'COALESCE(asm.tokensSeen, s.tokens_seen)');
  return db
    .prepare(
      `
    ${ACTIVE_SESSION_METRICS_CTE}
    SELECT
      s.video_id AS videoId,
      v.canonical_title AS canonicalTitle,
      a.canonical_title AS animeTitle,
      COALESCE(asm.activeWatchedMs, s.active_watched_ms, 0) AS activeWatchedMs,
      ${wordsExpr} AS tokensSeen,
      COALESCE(asm.cardsMined, s.cards_mined, 0) AS cardsMined,
      COALESCE(asm.yomitanLookupCount, s.yomitan_lookup_count, 0) AS yomitanLookupCount,
      CASE WHEN s.ended_at_ms IS NOT NULL THEN 1 ELSE 0 END AS ended,
      ${sessionRollupDaySql('s')} AS epochDay,
      ${sessionRollupMonthSql('s')} AS monthKey,
      CAST(strftime('%w', CAST(s.started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') AS INTEGER) AS dayOfWeek,
      CAST(strftime('%H', CAST(s.started_at_ms AS REAL) / 1000, 'unixepoch', 'localtime') AS INTEGER) AS hourOfDay
    FROM imm_sessions s
    LEFT JOIN active_session_metrics asm ON asm.sessionId = s.session_id
    LEFT JOIN imm_videos v ON v.video_id = s.video_id
    LEFT JOIN imm_anime a ON a.anime_id = v.anime_id
    WHERE ${sessionRollupDaySql('s')} BETWEEN ? AND ?
    ORDER BY s.started_at_ms ASC
  `,
    )
    .all(startDay ?? Number.MIN_SAFE_INTEGER, endDay) as unknown as TrendSessionRow[];
}

/**
 * Finished videos and the last day each was watched. `imm_videos.watched` has no
 * timestamp, so a video counts once, on the day of its last recorded activity.
 */
function getFinishedVideoDays(db: DatabaseSync): Array<{ videoId: number; epochDay: number }> {
  return db
    .prepare(
      `
        SELECT r.video_id AS videoId, MAX(r.rollup_day) AS epochDay
        FROM imm_daily_rollups r
        JOIN imm_videos v ON v.video_id = r.video_id
        WHERE v.watched = 1
        GROUP BY r.video_id
      `,
    )
    .all() as Array<{ videoId: number; epochDay: number }>;
}

/**
 * Rows behind the chart buckets. Day and week buckets read daily rollups. Month
 * buckets prefer monthly rollups, which outlive daily ones under the shorter
 * retention presets, but a month the range only partly covers has to come from
 * daily rows or it would include days outside the range.
 */
function bucketChartRollups(
  window: TrendWindow,
  groupBy: StatsTrendGroupBy,
  dailyRows: ImmersionSessionRollupRow[],
  monthlyRows: ImmersionSessionRollupRow[],
): Array<{ key: number; row: ImmersionSessionRollupRow }> {
  const { startDay, endDay, today } = window;
  if (groupBy !== 'month') {
    return dailyRows
      .filter((row) => startDay === null || row.rollupDayOrMonth >= startDay)
      .map((row) => ({ key: bucketKeyForDay(row.rollupDayOrMonth, groupBy), row }));
  }
  const isWholeMonth = (monthKey: number): boolean => {
    const bounds = monthDayBounds(monthKey);
    return (
      (startDay === null || bounds.startDay >= startDay) &&
      (bounds.endDay <= endDay || endDay >= today)
    );
  };
  return [
    ...monthlyRows
      .filter((row) => isWholeMonth(row.rollupDayOrMonth))
      .map((row) => ({ key: row.rollupDayOrMonth, row })),
    ...dailyRows
      .filter((row) => {
        const day = row.rollupDayOrMonth;
        return (startDay === null || day >= startDay) && !isWholeMonth(dayToMonthKey(day));
      })
      .map((row) => ({ key: dayToMonthKey(row.rollupDayOrMonth), row })),
  ];
}

function accumulatePoints(points: TrendChartPoint[]): TrendChartPoint[] {
  let sum = 0;
  return points.map((point) => {
    sum += point.value;
    return { label: point.label, value: sum };
  });
}

function toRate(numerator: number, denominator: number, scale: number, minDenominator: number) {
  return denominator >= minDenominator && denominator > 0
    ? +((numerator / denominator) * scale).toFixed(1)
    : null;
}

function buildLibrarySummary(
  rollups: ImmersionSessionRollupRow[],
  sessions: TrendSessionRow[],
  titlesByVideoId: Map<number, TitleRef>,
  finishedInWindow: Array<{ videoId: number }>,
): LibrarySummaryRow[] {
  type Accum = {
    animeId: number | null;
    latestVideo: { videoId: number; day: number } | null;
    watchTimeMin: number;
    videos: Set<number>;
    episodesFinished: number;
    cards: number;
    words: number;
    firstWatched: number;
    lastWatched: number;
    sessions: number;
    lookups: number;
  };

  const byTitle = new Map<string, Accum>();

  for (const rollup of rollups) {
    if (rollup.videoId === null) continue;
    const ref = titlesByVideoId.get(rollup.videoId);
    const title = ref?.title ?? 'Unknown';
    const acc: Accum = byTitle.get(title) ?? {
      animeId: ref?.animeId ?? null,
      latestVideo: null,
      watchTimeMin: 0,
      videos: new Set<number>(),
      episodesFinished: 0,
      cards: 0,
      words: 0,
      firstWatched: Number.POSITIVE_INFINITY,
      lastWatched: Number.NEGATIVE_INFINITY,
      sessions: 0,
      lookups: 0,
    };
    byTitle.set(title, acc);
    acc.watchTimeMin += rollup.totalActiveMin;
    acc.cards += rollup.totalCards;
    acc.words += rollup.totalTokensSeen;
    acc.videos.add(rollup.videoId);
    acc.firstWatched = Math.min(acc.firstWatched, rollup.rollupDayOrMonth);
    acc.lastWatched = Math.max(acc.lastWatched, rollup.rollupDayOrMonth);
    if (!acc.latestVideo || rollup.rollupDayOrMonth > acc.latestVideo.day) {
      acc.latestVideo = { videoId: rollup.videoId, day: rollup.rollupDayOrMonth };
    }
  }

  for (const session of sessions) {
    const acc = byTitle.get(resolveSessionTitle(session));
    if (!acc) continue;
    acc.sessions += 1;
    acc.lookups += session.yomitanLookupCount;
  }

  for (const { videoId } of finishedInWindow) {
    const title = titlesByVideoId.get(videoId)?.title;
    const acc = title === undefined ? undefined : byTitle.get(title);
    if (acc) acc.episodesFinished += 1;
  }

  const rows: LibrarySummaryRow[] = [];
  for (const [title, acc] of byTitle) {
    rows.push({
      title,
      animeId: acc.animeId,
      videoId: acc.latestVideo?.videoId ?? null,
      watchTimeMin: Math.round(acc.watchTimeMin),
      videos: acc.videos.size,
      episodesFinished: acc.episodesFinished,
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

function buildTitleTimeline(
  rollups: ImmersionSessionRollupRow[],
  titlesByVideoId: Map<number, TitleRef>,
): TrendPerAnimePoint[] {
  const byPoint = new Map<string, TrendPerAnimePoint>();
  for (const rollup of rollups) {
    if (rollup.videoId === null || rollup.totalActiveMin <= 0) continue;
    const animeTitle = titlesByVideoId.get(rollup.videoId)?.title ?? 'Unknown';
    const key = `${rollup.rollupDayOrMonth}:${animeTitle}`;
    const point = byPoint.get(key) ?? {
      epochDay: rollup.rollupDayOrMonth,
      animeTitle,
      value: 0,
    };
    point.value += rollup.totalActiveMin;
    byPoint.set(key, point);
  }
  return [...byPoint.values()].map((point) => ({
    ...point,
    value: Math.round(point.value * 10) / 10,
  }));
}

function buildWatchTimeByDayHour(sessions: TrendSessionRow[]): number[][] {
  const grid = Array.from({ length: 7 }, () => new Array<number>(24).fill(0));
  for (const session of sessions) {
    grid[session.dayOfWeek]![session.hourOfDay]! += session.activeWatchedMs;
  }
  return grid.map((hours) => hours.map((ms) => Math.round(ms / 60_000)));
}

function buildFrequencyBands(
  rows: NewWordFirstSeenRow[],
  window: { startDay: number; endDay: number },
  groupBy: StatsTrendGroupBy,
): TrendFrequencyBandPoint[] {
  // A band mix needs more words than most single days have, so daily views use weeks.
  const bandGroupBy = groupBy === 'day' ? 'week' : groupBy;
  const countsByKey = new Map<number, TrendFrequencyBandPoint['counts']>();
  for (const row of rows) {
    const day = row.firstSeenDayWithoutNames;
    if (day === null || day < window.startDay || day > window.endDay) continue;
    const key = bucketKeyForDay(day, bandGroupBy);
    const counts =
      countsByKey.get(key) ??
      (Object.fromEntries(
        TREND_FREQUENCY_BANDS.map((band) => [band, 0]),
      ) as TrendFrequencyBandPoint['counts']);
    counts[row.band] += row.wordCount;
    countsByKey.set(key, counts);
  }
  return buildBuckets(window.startDay, window.endDay, bandGroupBy, new Set(countsByKey.keys())).map(
    (bucket) => ({ label: bucket.label, counts: countsByKey.get(bucket.key)! }),
  );
}

/**
 * Builds every series on the Trends tab for a date range. A `title` narrows the
 * totals, charts, and patterns to that title; `librarySummary` and `titleTimeline`
 * stay complete so the title can still be compared to and switched from the rest.
 */
export function getTrendsDashboard(
  db: DatabaseSync,
  range: StatsTrendRange = '30d',
  groupBy: StatsTrendGroupBy = 'day',
  fillEmptyBuckets = true,
  title: string | null = null,
): TrendsDashboardData {
  const window = resolveTrendWindow(db, range);
  const { startDay, endDay, previous, loadFromDay } = window;

  const sessionRows = getTrendSessionRows(db, loadFromDay, endDay);
  const dailyWordCounts: RollupWordCounts = new Map();
  const monthlyWordCounts: RollupWordCounts = new Map();
  for (const row of sessionRows) {
    if (row.ended !== 1) continue;
    addRollupWordCount(dailyWordCounts, row.epochDay, row.videoId, row.tokensSeen);
    addRollupWordCount(monthlyWordCounts, row.monthKey, row.videoId, row.tokensSeen);
  }
  const allDailyRollups = applyRollupWordCounts(
    selectDailyRollupRowsBetween(db, loadFromDay, endDay),
    dailyWordCounts,
  );
  const monthlyRollups =
    groupBy === 'month'
      ? applyRollupWordCounts(
          selectMonthlyRollupRowsBetween(
            db,
            startDay === null ? null : dayToMonthKey(startDay),
            dayToMonthKey(endDay),
          ),
          monthlyWordCounts,
        )
      : [];
  const finishedVideos = getFinishedVideoDays(db);
  const titlesByVideoId = getVideoTitleMap(db, [
    ...allDailyRollups.map((row) => row.videoId),
    ...monthlyRollups.map((row) => row.videoId),
    ...finishedVideos.map((row) => row.videoId),
  ]);

  const inWindow = (day: number): boolean =>
    (startDay === null || day >= startDay) && day <= endDay;
  const matchesTitle = (videoId: number | null): boolean =>
    title === null || (videoId !== null && titlesByVideoId.get(videoId)?.title === title);

  // New words are recorded per word, not per title, so a title filter has none.
  const newWordRows = title === null ? getNewWordFirstSeenRows(db) : [];

  const dailyRollups = allDailyRollups.filter((row) => matchesTitle(row.videoId));
  const sessions =
    title === null ? sessionRows : sessionRows.filter((row) => resolveSessionTitle(row) === title);

  const totalsByDay = new Map<number, TrendTotals>();
  const dayTotals = (day: number): TrendTotals => {
    const totals = totalsByDay.get(day) ?? emptyTotals();
    totalsByDay.set(day, totals);
    return totals;
  };
  const rolledUpDays = new Set<number>();
  for (const row of dailyRollups) {
    const totals = dayTotals(row.rollupDayOrMonth);
    totals.activeMin += row.totalActiveMin;
    totals.cards += row.totalCards;
    totals.words += row.totalTokensSeen;
    rolledUpDays.add(row.rollupDayOrMonth);
  }
  for (const session of sessions) {
    if (session.ended === 1) dayTotals(session.epochDay).lookups += session.yomitanLookupCount;
  }
  for (const finished of finishedVideos) {
    if (matchesTitle(finished.videoId)) dayTotals(finished.epochDay).episodes += 1;
  }
  for (const row of newWordRows) {
    if (row.firstSeenDayWithoutNames !== null) {
      dayTotals(row.firstSeenDayWithoutNames).newWords += row.wordCount;
    }
  }

  // Bucket totals: rollup metrics come from the rows the bucket should read, while
  // everything that only exists per day folds up from `totalsByDay`.
  const chartRollups = bucketChartRollups(
    window,
    groupBy,
    dailyRollups,
    monthlyRollups.filter((row) => matchesTitle(row.videoId)),
  );
  const totalsByBucket = new Map<number, TrendTotals>();
  const bucketTotals = (key: number): TrendTotals => {
    const totals = totalsByBucket.get(key) ?? emptyTotals();
    totalsByBucket.set(key, totals);
    return totals;
  };
  for (const { key, row } of chartRollups) {
    const totals = bucketTotals(key);
    totals.activeMin += row.totalActiveMin;
    totals.cards += row.totalCards;
    totals.words += row.totalTokensSeen;
  }
  for (const [day, totals] of totalsByDay) {
    if (!inWindow(day)) continue;
    const target = bucketTotals(bucketKeyForDay(day, groupBy));
    target.lookups += totals.lookups;
    target.episodes += totals.episodes;
    target.newWords += totals.newWords;
  }

  // `all` spans the recorded data, including months only monthly rollups still hold.
  const activeKeys = new Set(totalsByBucket.keys());
  const activeDays = [...totalsByDay.keys()].filter(inWindow);
  const monthlyKeys = monthlyRollups.map((row) => row.rollupDayOrMonth);
  const dataDays = [
    ...activeDays,
    ...(monthlyKeys.length > 0
      ? [
          monthDayBounds(Math.min(...monthlyKeys)).startDay,
          Math.min(monthDayBounds(Math.max(...monthlyKeys)).endDay, endDay),
        ]
      : []),
  ];
  const dataStart = dataDays.length > 0 ? Math.min(...dataDays) : null;
  const dataEnd = dataDays.length > 0 ? Math.max(...dataDays) : null;
  const axisStart = startDay ?? dataStart;
  const axisEnd = startDay === null ? dataEnd : endDay;
  const buckets =
    axisStart === null || axisEnd === null
      ? []
      : buildBuckets(axisStart, axisEnd, groupBy, fillEmptyBuckets ? null : activeKeys);

  const series = (pick: (totals: TrendTotals) => number): TrendChartPoint[] =>
    buckets.map((bucket) => {
      const totals = totalsByBucket.get(bucket.key);
      return { label: bucket.label, value: totals ? pick(totals) : 0 };
    });

  const rateSeries = (
    numerator: (totals: TrendTotals) => number,
    denominator: (totals: TrendTotals) => number,
    scale: number,
    minDenominator: number,
  ): TrendRatePoint[] =>
    buckets.map((bucket) => {
      if (groupBy !== 'day') {
        const totals = totalsByBucket.get(bucket.key) ?? emptyTotals();
        return {
          label: bucket.label,
          value: toRate(numerator(totals), denominator(totals), scale, minDenominator),
        };
      }
      let top = 0;
      let bottom = 0;
      for (let day = bucket.key - (ROLLING_RATE_DAYS - 1); day <= bucket.key; day += 1) {
        const totals = totalsByDay.get(day);
        if (!totals) continue;
        top += numerator(totals);
        bottom += denominator(totals);
      }
      return { label: bucket.label, value: toRate(top, bottom, scale, minDenominator) };
    });

  const sumDays = (span: DayRange): TrendSummaryTotals => {
    const sum = emptyTotals();
    let activeDayCount = 0;
    for (const [day, totals] of totalsByDay) {
      if (day < span.startDay || day > span.endDay) continue;
      sum.activeMin += totals.activeMin;
      sum.cards += totals.cards;
      sum.words += totals.words;
      sum.episodes += totals.episodes;
      sum.newWords += totals.newWords;
      if (rolledUpDays.has(day)) activeDayCount += 1;
    }
    return {
      watchTimeMin: Math.round(sum.activeMin),
      words: sum.words,
      newWords: title === null ? sum.newWords : null,
      cards: sum.cards,
      episodes: sum.episodes,
      activeDays: activeDayCount,
    };
  };

  // In date order like the running totals: float addition depends on order, so a total
  // near .5 could otherwise round differently here than where the chart ends.
  const bucketSum = [...totalsByBucket.entries()]
    .sort(([left], [right]) => left - right)
    .reduce<TrendTotals>(
      (sum, [, totals]) => ({
        activeMin: sum.activeMin + totals.activeMin,
        cards: sum.cards + totals.cards,
        words: sum.words + totals.words,
        lookups: sum.lookups + totals.lookups,
        episodes: sum.episodes + totals.episodes,
        newWords: sum.newWords + totals.newWords,
      }),
      emptyTotals(),
    );

  const windowRollups = allDailyRollups.filter((row) => inWindow(row.rollupDayOrMonth));
  const windowSessions = sessionRows.filter((row) => inWindow(row.epochDay));
  const reportedWindow = { startDay: axisStart ?? endDay, endDay: axisEnd ?? endDay };

  const activity = {
    watchTime: series((totals) => Math.round(totals.activeMin)),
    cards: series((totals) => totals.cards),
    words: series((totals) => totals.words),
    episodes: series((totals) => totals.episodes),
  };

  return {
    window: reportedWindow,
    buckets: buckets.map(({ label, startDay: first, endDay: last }) => ({
      label,
      startDay: first,
      endDay: last,
    })),
    summary: {
      current: {
        watchTimeMin: Math.round(bucketSum.activeMin),
        words: bucketSum.words,
        newWords: title === null ? bucketSum.newWords : null,
        cards: bucketSum.cards,
        episodes: bucketSum.episodes,
        activeDays: [...rolledUpDays].filter(inWindow).length,
      },
      previous: previous ? sumDays(previous) : null,
    },
    activity,
    progress: {
      // Rounded after summing, so rounding each day's minutes can't drift from the summary total.
      watchTime: accumulatePoints(series((totals) => totals.activeMin)).map((point) => ({
        ...point,
        value: Math.round(point.value),
      })),
      words: accumulatePoints(activity.words),
      newWords: title === null ? accumulatePoints(series((totals) => totals.newWords)) : [],
      cards: accumulatePoints(activity.cards),
      episodes: accumulatePoints(activity.episodes),
      lookups: accumulatePoints(series((totals) => totals.lookups)),
    },
    ratios: {
      lookupsPerHundred: rateSeries(
        (totals) => totals.lookups,
        (totals) => totals.words,
        100,
        MIN_RATE_WORDS,
      ),
      cardsPerHour: rateSeries(
        (totals) => totals.cards,
        (totals) => totals.activeMin,
        60,
        MIN_RATE_ACTIVE_MIN,
      ),
      readingSpeed: rateSeries(
        (totals) => totals.words,
        (totals) => totals.activeMin,
        1,
        MIN_RATE_ACTIVE_MIN,
      ),
    },
    learning:
      title === null
        ? {
            newWordsPerThousand: rateSeries(
              (totals) => totals.newWords,
              (totals) => totals.words,
              1_000,
              MIN_RATE_WORDS,
            ),
            frequencyBands: buildFrequencyBands(newWordRows, reportedWindow, groupBy),
          }
        : null,
    watchTimeByDayHour: buildWatchTimeByDayHour(sessions.filter((row) => inWindow(row.epochDay))),
    librarySummary: buildLibrarySummary(
      windowRollups,
      windowSessions,
      titlesByVideoId,
      finishedVideos.filter((row) => inWindow(row.epochDay)),
    ),
    titleTimeline: buildTitleTimeline(windowRollups, titlesByVideoId),
  };
}
