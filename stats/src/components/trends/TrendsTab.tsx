import { useState } from 'react';
import { useTrends, type TimeRange, type GroupBy } from '../../hooks/useTrends';
import { DateRangeSelector } from './DateRangeSelector';
import { TrendChart } from './TrendChart';
import { StackedTrendChart, type PerAnimeDataPoint } from './StackedTrendChart';
import {
  buildAnimeVisibilityOptions,
  filterHiddenAnimeData,
  pruneHiddenAnime,
} from './anime-visibility';
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

interface AnimeVisibilityFilterProps {
  animeTitles: string[];
  hiddenAnime: ReadonlySet<string>;
  onShowAll: () => void;
  onHideAll: () => void;
  onToggleAnime: (title: string) => void;
}

function AnimeVisibilityFilter({
  animeTitles,
  hiddenAnime,
  onShowAll,
  onHideAll,
  onToggleAnime,
}: AnimeVisibilityFilterProps) {
  if (animeTitles.length === 0) {
    return null;
  }

  return (
    <div className="col-span-full -mt-1 mb-1 rounded-lg border border-ctp-surface1 bg-ctp-surface0/70 p-3">
      <div className="mb-2 flex items-center justify-between gap-3">
        <div>
          <h4 className="text-xs font-semibold uppercase tracking-widest text-ctp-subtext0">
            Anime Visibility
          </h4>
          <p className="mt-1 text-xs text-ctp-overlay1">
            Shared across all anime trend charts. Default: show everything.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            className="rounded-md border border-ctp-surface2 px-2 py-1 text-[11px] font-medium text-ctp-text transition hover:border-ctp-blue hover:text-ctp-blue"
            onClick={onShowAll}
          >
            All
          </button>
          <button
            type="button"
            className="rounded-md border border-ctp-surface2 px-2 py-1 text-[11px] font-medium text-ctp-text transition hover:border-ctp-peach hover:text-ctp-peach"
            onClick={onHideAll}
          >
            None
          </button>
        </div>
      </div>
      <div className="flex flex-wrap gap-2">
        {animeTitles.map((title) => {
          const isVisible = !hiddenAnime.has(title);
          return (
            <button
              key={title}
              type="button"
              aria-pressed={isVisible}
              className={`max-w-full rounded-full border px-3 py-1 text-xs transition ${
                isVisible
                  ? 'border-ctp-blue/60 bg-ctp-blue/12 text-ctp-blue'
                  : 'border-ctp-surface2 bg-transparent text-ctp-subtext0'
              }`}
              onClick={() => onToggleAnime(title)}
              title={title}
            >
              <span className="block truncate">{title}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

export function TrendsTab() {
  const [range, setRange] = useState<TimeRange>('30d');
  const [groupBy, setGroupBy] = useState<GroupBy>('day');
  const [hiddenAnime, setHiddenAnime] = useState<Set<string>>(() => new Set());
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
  const animeTitles = buildAnimeVisibilityOptions([
    episodesPerAnime,
    watchTimePerAnime,
    cardsPerAnime,
    wordsPerAnime,
    animeProgress,
    cardsProgress,
    wordsProgress,
  ]);
  const activeHiddenAnime = pruneHiddenAnime(hiddenAnime, animeTitles);

  const filteredEpisodesPerAnime = filterHiddenAnimeData(episodesPerAnime, activeHiddenAnime);
  const filteredWatchTimePerAnime = filterHiddenAnimeData(watchTimePerAnime, activeHiddenAnime);
  const filteredCardsPerAnime = filterHiddenAnimeData(cardsPerAnime, activeHiddenAnime);
  const filteredWordsPerAnime = filterHiddenAnimeData(wordsPerAnime, activeHiddenAnime);
  const filteredAnimeProgress = filterHiddenAnimeData(animeProgress, activeHiddenAnime);
  const filteredCardsProgress = filterHiddenAnimeData(cardsProgress, activeHiddenAnime);
  const filteredWordsProgress = filterHiddenAnimeData(wordsProgress, activeHiddenAnime);

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
        <AnimeVisibilityFilter
          animeTitles={animeTitles}
          hiddenAnime={activeHiddenAnime}
          onShowAll={() => setHiddenAnime(new Set())}
          onHideAll={() => setHiddenAnime(new Set(animeTitles))}
          onToggleAnime={(title) =>
            setHiddenAnime((current) => {
              const next = new Set(current);
              if (next.has(title)) {
                next.delete(title);
              } else {
                next.add(title);
              }
              return next;
            })
          }
        />
        <StackedTrendChart title="Episodes per Anime" data={filteredEpisodesPerAnime} />
        <StackedTrendChart title="Watch Time per Anime (min)" data={filteredWatchTimePerAnime} />
        <StackedTrendChart title="Cards Mined per Anime" data={filteredCardsPerAnime} />
        <StackedTrendChart title="Words Seen per Anime" data={filteredWordsPerAnime} />

        <SectionHeader>Anime — Cumulative</SectionHeader>
        <StackedTrendChart title="Episodes Progress" data={filteredAnimeProgress} />
        <StackedTrendChart title="Cards Mined Progress" data={filteredCardsProgress} />
        <StackedTrendChart title="Words Seen Progress" data={filteredWordsProgress} />

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
