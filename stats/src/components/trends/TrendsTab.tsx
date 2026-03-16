import { useState } from 'react';
import { useTrends, type TimeRange, type GroupBy } from '../../hooks/useTrends';
import { DateRangeSelector } from './DateRangeSelector';
import { TrendChart } from './TrendChart';
import { StackedTrendChart, type PerAnimeDataPoint } from './StackedTrendChart';
import { buildTrendDashboard, type ChartPoint } from '../../lib/dashboard-data';
import { localDayFromMs } from '../../lib/formatters';
import type { SessionSummary } from '../../types/stats';

const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

function buildWatchTimeByDayOfWeek(sessions: SessionSummary[]): ChartPoint[] {
  const totals = new Array(7).fill(0);
  for (const s of sessions) {
    const dow = new Date(s.startedAtMs).getDay();
    totals[dow] += s.activeWatchedMs;
  }
  return DAY_NAMES.map((name, i) => ({ label: name, value: Math.round(totals[i] / 60_000) }));
}

function buildWatchTimeByHour(sessions: SessionSummary[]): ChartPoint[] {
  const totals = new Array(24).fill(0);
  for (const s of sessions) {
    const hour = new Date(s.startedAtMs).getHours();
    totals[hour] += s.activeWatchedMs;
  }
  return totals.map((ms, i) => ({
    label: `${String(i).padStart(2, '0')}:00`,
    value: Math.round(ms / 60_000),
  }));
}

function buildCumulativePerAnime(points: PerAnimeDataPoint[]): PerAnimeDataPoint[] {
  const byAnime = new Map<string, Map<number, number>>();
  const allDays = new Set<number>();
  for (const p of points) {
    const dayMap = byAnime.get(p.animeTitle) ?? new Map();
    dayMap.set(p.epochDay, (dayMap.get(p.epochDay) ?? 0) + p.value);
    byAnime.set(p.animeTitle, dayMap);
    allDays.add(p.epochDay);
  }

  const sortedDays = [...allDays].sort((a, b) => a - b);
  if (sortedDays.length < 2) return points;

  const minDay = sortedDays[0]!;
  const maxDay = sortedDays[sortedDays.length - 1]!;
  const everyDay: number[] = [];
  for (let d = minDay; d <= maxDay; d++) {
    everyDay.push(d);
  }

  const result: PerAnimeDataPoint[] = [];
  for (const [animeTitle, dayMap] of byAnime) {
    let cumulative = 0;
    const firstDay = Math.min(...dayMap.keys());
    for (const day of everyDay) {
      if (day < firstDay) continue;
      cumulative += dayMap.get(day) ?? 0;
      result.push({ epochDay: day, animeTitle, value: cumulative });
    }
  }
  return result;
}

function buildPerAnimeFromSessions(
  sessions: SessionSummary[],
  getValue: (s: SessionSummary) => number,
): PerAnimeDataPoint[] {
  const map = new Map<string, Map<number, number>>();
  for (const s of sessions) {
    const title = s.animeTitle ?? s.canonicalTitle ?? 'Unknown';
    const day = localDayFromMs(s.startedAtMs);
    const animeMap = map.get(title) ?? new Map();
    animeMap.set(day, (animeMap.get(day) ?? 0) + getValue(s));
    map.set(title, animeMap);
  }
  const points: PerAnimeDataPoint[] = [];
  for (const [animeTitle, dayMap] of map) {
    for (const [epochDay, value] of dayMap) {
      points.push({ epochDay, animeTitle, value });
    }
  }
  return points;
}

function buildEpisodesPerAnimeFromSessions(sessions: SessionSummary[]): PerAnimeDataPoint[] {
  // Group by anime+day, counting distinct videoIds
  const map = new Map<string, Map<number, Set<number | null>>>();
  for (const s of sessions) {
    const title = s.animeTitle ?? s.canonicalTitle ?? 'Unknown';
    const day = localDayFromMs(s.startedAtMs);
    const animeMap = map.get(title) ?? new Map();
    const videoSet = animeMap.get(day) ?? new Set();
    videoSet.add(s.videoId);
    animeMap.set(day, videoSet);
    map.set(title, animeMap);
  }
  const points: PerAnimeDataPoint[] = [];
  for (const [animeTitle, dayMap] of map) {
    for (const [epochDay, videoSet] of dayMap) {
      points.push({ epochDay, animeTitle, value: videoSet.size });
    }
  }
  return points;
}

function SectionHeader({ children }: { children: React.ReactNode }) {
  return (
    <div className="col-span-full mt-6 mb-2 flex items-center gap-3">
      <h3 className="text-ctp-subtext0 text-xs font-semibold uppercase tracking-widest shrink-0">
        {children}
      </h3>
      <div className="flex-1 h-px bg-gradient-to-r from-ctp-surface1 to-transparent" />
    </div>
  );
}

export function TrendsTab() {
  const [range, setRange] = useState<TimeRange>('30d');
  const [groupBy, setGroupBy] = useState<GroupBy>('day');
  const { data, loading, error } = useTrends(range, groupBy);

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;

  const dashboard = buildTrendDashboard(data.rollups);
  const watchByDow = buildWatchTimeByDayOfWeek(data.sessions);
  const watchByHour = buildWatchTimeByHour(data.sessions);

  const watchTimePerAnime = data.watchTimePerAnime.map((e) => ({
    epochDay: e.epochDay,
    animeTitle: e.animeTitle,
    value: e.totalActiveMin,
  }));
  const episodesPerAnime = buildEpisodesPerAnimeFromSessions(data.sessions);
  const cardsPerAnime = buildPerAnimeFromSessions(data.sessions, (s) => s.cardsMined);
  const wordsPerAnime = buildPerAnimeFromSessions(data.sessions, (s) => s.wordsSeen);

  const animeProgress = buildCumulativePerAnime(episodesPerAnime);
  const cardsProgress = buildCumulativePerAnime(cardsPerAnime);
  const wordsProgress = buildCumulativePerAnime(wordsPerAnime);

  return (
    <div className="space-y-4">
      <DateRangeSelector
        range={range}
        groupBy={groupBy}
        onRangeChange={setRange}
        onGroupByChange={setGroupBy}
      />
      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
        <SectionHeader>Activity</SectionHeader>
        <TrendChart
          title="Watch Time (min)"
          data={dashboard.watchTime}
          color="#8aadf4"
          type="bar"
        />
        <TrendChart title="Cards Mined" data={dashboard.cards} color="#a6da95" type="bar" />
        <TrendChart title="Words Seen" data={dashboard.words} color="#8bd5ca" type="bar" />
        <TrendChart title="Sessions" data={dashboard.sessions} color="#b7bdf8" type="line" />
        <TrendChart
          title="Avg Session (min)"
          data={dashboard.averageSessionMinutes}
          color="#f5bde6"
          type="line"
        />

        <SectionHeader>Anime — Per Day</SectionHeader>
        <StackedTrendChart title="Episodes per Anime" data={episodesPerAnime} />
        <StackedTrendChart title="Watch Time per Anime (min)" data={watchTimePerAnime} />
        <StackedTrendChart title="Cards Mined per Anime" data={cardsPerAnime} />
        <StackedTrendChart title="Words Seen per Anime" data={wordsPerAnime} />

        <SectionHeader>Anime — Cumulative</SectionHeader>
        <StackedTrendChart title="Episodes Progress" data={animeProgress} />
        <StackedTrendChart title="Cards Mined Progress" data={cardsProgress} />
        <StackedTrendChart title="Words Seen Progress" data={wordsProgress} />

        <SectionHeader>Patterns</SectionHeader>
        <TrendChart
          title="Watch Time by Day of Week (min)"
          data={watchByDow}
          color="#8aadf4"
          type="bar"
        />
        <TrendChart
          title="Watch Time by Hour (min)"
          data={watchByHour}
          color="#c6a0f6"
          type="bar"
        />
      </div>
    </div>
  );
}
