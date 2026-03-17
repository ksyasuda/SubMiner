import type {
  DailyRollup,
  KanjiEntry,
  OverviewData,
  StreakCalendarDay,
  VocabularyEntry,
} from '../types/stats';
import { epochDayToDate, epochMsFromDbTimestamp, localDayFromMs } from './formatters';

export interface ChartPoint {
  label: string;
  value: number;
}

export interface OverviewSummary {
  todayActiveMs: number;
  todayCards: number;
  streakDays: number;
  allTimeHours: number;
  totalTrackedCards: number;
  episodesToday: number;
  activeAnimeCount: number;
  totalEpisodesWatched: number;
  totalAnimeCompleted: number;
  averageSessionMinutes: number;
  totalSessions: number;
  activeDays: number;
  recentWatchTime: ChartPoint[];
}

export interface TrendDashboard {
  watchTime: ChartPoint[];
  cards: ChartPoint[];
  words: ChartPoint[];
  sessions: ChartPoint[];
  cardsPerHour: ChartPoint[];
  lookupHitRate: ChartPoint[];
  averageSessionMinutes: ChartPoint[];
}

export interface VocabularySummary {
  uniqueWords: number;
  uniqueKanji: number;
  newThisWeek: number;
  topWords: ChartPoint[];
  newWordsTimeline: ChartPoint[];
  recentDiscoveries: VocabularyEntry[];
}

function normalizeDbTimestampSeconds(ts: number): number {
  return Math.floor(epochMsFromDbTimestamp(ts) / 1000);
}

function makeRollupLabel(value: number): string {
  if (value > 100_000) {
    const year = Math.floor(value / 100);
    const month = value % 100;
    return new Date(Date.UTC(year, month - 1, 1)).toLocaleDateString(undefined, {
      month: 'short',
      year: '2-digit',
    });
  }

  return epochDayToDate(value).toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
  });
}

function sumBy<T>(values: T[], select: (value: T) => number): number {
  return values.reduce((sum, value) => sum + select(value), 0);
}

function buildAggregatedDailyRows(rollups: DailyRollup[]) {
  const byKey = new Map<
    number,
    {
      activeMin: number;
      cards: number;
      words: number;
      sessions: number;
      lookupHitRateSum: number;
      lookupWeight: number;
    }
  >();

  for (const rollup of rollups) {
    const existing = byKey.get(rollup.rollupDayOrMonth) ?? {
      activeMin: 0,
      cards: 0,
      words: 0,
      sessions: 0,
      lookupHitRateSum: 0,
      lookupWeight: 0,
    };

    existing.activeMin += rollup.totalActiveMin;
    existing.cards += rollup.totalCards;
    existing.words += rollup.totalWordsSeen;
    existing.sessions += rollup.totalSessions;
    if (rollup.lookupHitRate != null) {
      const weight = Math.max(rollup.totalSessions, 1);
      existing.lookupHitRateSum += rollup.lookupHitRate * weight;
      existing.lookupWeight += weight;
    }

    byKey.set(rollup.rollupDayOrMonth, existing);
  }

  return Array.from(byKey.entries())
    .sort(([left], [right]) => left - right)
    .map(([key, value]) => ({
      key,
      label: makeRollupLabel(key),
      activeMin: Math.round(value.activeMin),
      cards: value.cards,
      words: value.words,
      sessions: value.sessions,
      cardsPerHour: value.activeMin > 0 ? +((value.cards * 60) / value.activeMin).toFixed(1) : 0,
      averageSessionMinutes:
        value.sessions > 0 ? +(value.activeMin / value.sessions).toFixed(1) : 0,
      lookupHitRate:
        value.lookupWeight > 0
          ? Math.round((value.lookupHitRateSum / value.lookupWeight) * 100)
          : 0,
    }));
}

export function buildOverviewSummary(
  overview: OverviewData,
  nowMs: number = Date.now(),
): OverviewSummary {
  const today = localDayFromMs(nowMs);
  const aggregated = buildAggregatedDailyRows(overview.rollups);
  const todayRow = aggregated.find((row) => row.key === today);
  const daysWithActivity = new Set(
    aggregated.filter((row) => row.activeMin > 0).map((row) => row.key),
  );

  const sessionCards = sumBy(overview.sessions, (session) => session.cardsMined);
  const rollupCards = sumBy(aggregated, (row) => row.cards);
  const lifetimeCards = overview.hints.totalCards ?? Math.max(sessionCards, rollupCards);
  const totalActiveMin = overview.hints.totalActiveMin ?? sumBy(aggregated, (row) => row.activeMin);

  let streakDays = 0;
  const streakStart = daysWithActivity.has(today) ? today : today - 1;
  for (let day = streakStart; daysWithActivity.has(day); day -= 1) {
    streakDays += 1;
  }

  const todaySessions = overview.sessions.filter(
    (session) => localDayFromMs(session.startedAtMs) === today,
  );
  const todayActiveFromSessions = sumBy(todaySessions, (session) => session.activeWatchedMs);
  const todayActiveFromRollup = (todayRow?.activeMin ?? 0) * 60_000;

  return {
    todayActiveMs: Math.max(todayActiveFromRollup, todayActiveFromSessions),
    todayCards: Math.max(
      todayRow?.cards ?? 0,
      sumBy(todaySessions, (session) => session.cardsMined),
    ),
    streakDays,
    allTimeHours: Math.max(0, Math.round(totalActiveMin / 60)),
    totalTrackedCards: lifetimeCards,
    episodesToday: overview.hints.episodesToday ?? 0,
    activeAnimeCount: overview.hints.activeAnimeCount ?? 0,
    totalEpisodesWatched: overview.hints.totalEpisodesWatched ?? 0,
    totalAnimeCompleted: overview.hints.totalAnimeCompleted ?? 0,
    averageSessionMinutes:
      overview.sessions.length > 0
        ? Math.round(
            sumBy(overview.sessions, (session) => session.activeWatchedMs) /
              overview.sessions.length /
              60_000,
          )
        : 0,
    totalSessions: overview.hints.totalSessions,
    activeDays: overview.hints.activeDays ?? daysWithActivity.size,
    recentWatchTime: aggregated
      .slice(-14)
      .map((row) => ({ label: row.label, value: row.activeMin })),
  };
}

export function buildTrendDashboard(rollups: DailyRollup[]): TrendDashboard {
  const aggregated = buildAggregatedDailyRows(rollups);
  return {
    watchTime: aggregated.map((row) => ({ label: row.label, value: row.activeMin })),
    cards: aggregated.map((row) => ({ label: row.label, value: row.cards })),
    words: aggregated.map((row) => ({ label: row.label, value: row.words })),
    sessions: aggregated.map((row) => ({ label: row.label, value: row.sessions })),
    cardsPerHour: aggregated.map((row) => ({ label: row.label, value: row.cardsPerHour })),
    lookupHitRate: aggregated.map((row) => ({ label: row.label, value: row.lookupHitRate })),
    averageSessionMinutes: aggregated.map((row) => ({
      label: row.label,
      value: row.averageSessionMinutes,
    })),
  };
}

export function buildVocabularySummary(
  words: VocabularyEntry[],
  kanji: KanjiEntry[],
  nowMs: number = Date.now(),
): VocabularySummary {
  const weekAgoSec = nowMs / 1000 - 7 * 86_400;
  const byDay = new Map<number, number>();

  for (const word of words) {
    const firstSeenSec = normalizeDbTimestampSeconds(word.firstSeen);
    const day = Math.floor(firstSeenSec / 86_400);
    byDay.set(day, (byDay.get(day) ?? 0) + 1);
  }

  return {
    uniqueWords: words.length,
    uniqueKanji: kanji.length,
    newThisWeek: words.filter((word) => {
      const firstSeenSec = normalizeDbTimestampSeconds(word.firstSeen);
      return firstSeenSec >= weekAgoSec;
    }).length,
    topWords: [...words]
      .sort((left, right) => right.frequency - left.frequency)
      .slice(0, 12)
      .map((word) => ({ label: word.headword, value: word.frequency })),
    newWordsTimeline: Array.from(byDay.entries())
      .sort(([left], [right]) => left - right)
      .slice(-14)
      .map(([day, count]) => ({
        label: makeRollupLabel(day),
        value: count,
      })),
    recentDiscoveries: [...words]
      .sort((left, right) => {
        const leftFirst = normalizeDbTimestampSeconds(left.firstSeen);
        const rightFirst = normalizeDbTimestampSeconds(right.firstSeen);
        return rightFirst - leftFirst;
      })
      .slice(0, 8),
  };
}

export interface StreakCalendarPoint {
  date: string;
  value: number;
}

export function buildStreakCalendar(days: StreakCalendarDay[]): StreakCalendarPoint[] {
  return days.map((d) => {
    const dt = epochDayToDate(d.epochDay);
    const y = dt.getUTCFullYear();
    const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const day = String(dt.getUTCDate()).padStart(2, '0');
    return { date: `${y}-${m}-${day}`, value: d.totalActiveMin };
  });
}
