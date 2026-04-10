import { useState } from 'react';
import { useTrends, type TimeRange, type GroupBy } from '../../hooks/useTrends';
import { DateRangeSelector } from './DateRangeSelector';
import { TrendChart } from './TrendChart';
import { StackedTrendChart } from './StackedTrendChart';
import {
  buildAnimeVisibilityOptions,
  filterHiddenAnimeData,
  pruneHiddenAnime,
} from './anime-visibility';
import { LibrarySummarySection } from './LibrarySummarySection';

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
  const cardsMinedColor = 'var(--color-ctp-cards-mined)';
  const cardsMinedStackedColors = [
    cardsMinedColor,
    '#8aadf4',
    '#c6a0f6',
    '#f5a97f',
    '#f5bde6',
    '#91d7e3',
    '#ee99a0',
    '#f4dbd6',
  ];

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data) return null;

  const librarySummaryAsPoints = data.librarySummary.map((row) => ({
    epochDay: 0,
    animeTitle: row.title,
    value: row.watchTimeMin,
  }));

  const animeTitles = buildAnimeVisibilityOptions([
    librarySummaryAsPoints,
    data.animeCumulative.episodes,
    data.animeCumulative.cards,
    data.animeCumulative.words,
    data.animeCumulative.watchTime,
  ]);
  const activeHiddenAnime = pruneHiddenAnime(hiddenAnime, animeTitles);

  const filteredAnimeProgress = filterHiddenAnimeData(
    data.animeCumulative.episodes,
    activeHiddenAnime,
  );
  const filteredCardsProgress = filterHiddenAnimeData(
    data.animeCumulative.cards,
    activeHiddenAnime,
  );
  const filteredWordsProgress = filterHiddenAnimeData(
    data.animeCumulative.words,
    activeHiddenAnime,
  );
  const filteredWatchTimeProgress = filterHiddenAnimeData(
    data.animeCumulative.watchTime,
    activeHiddenAnime,
  );

  return (
    <div className="space-y-4">
      <DateRangeSelector
        range={range}
        groupBy={groupBy}
        onRangeChange={setRange}
        onGroupByChange={setGroupBy}
      />
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <SectionHeader>Activity</SectionHeader>
        <TrendChart
          title="Watch Time (min)"
          data={data.activity.watchTime}
          color="#8aadf4"
          type="bar"
        />
        <TrendChart
          title="Cards Mined"
          data={data.activity.cards}
          color={cardsMinedColor}
          type="bar"
        />
        <TrendChart title="Words Seen" data={data.activity.words} color="#8bd5ca" type="bar" />
        <TrendChart title="Sessions" data={data.activity.sessions} color="#b7bdf8" type="bar" />

        <SectionHeader>Period Trends</SectionHeader>
        <TrendChart
          title="Watch Time (min)"
          data={data.progress.watchTime}
          color="#8aadf4"
          type="line"
        />
        <TrendChart title="Sessions" data={data.progress.sessions} color="#b7bdf8" type="line" />
        <TrendChart title="Words Seen" data={data.progress.words} color="#8bd5ca" type="line" />
        <TrendChart
          title="New Words Seen"
          data={data.progress.newWords}
          color="#c6a0f6"
          type="line"
        />
        <TrendChart
          title="Cards Mined"
          data={data.progress.cards}
          color={cardsMinedColor}
          type="line"
        />
        <TrendChart
          title="Episodes Watched"
          data={data.progress.episodes}
          color="#91d7e3"
          type="line"
        />
        <TrendChart title="Lookups" data={data.progress.lookups} color="#f5bde6" type="line" />
        <TrendChart
          title="Lookups / 100 Words"
          data={data.ratios.lookupsPerHundred}
          color="#f5a97f"
          type="line"
        />

        <SectionHeader>Library — Summary</SectionHeader>
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
        <LibrarySummarySection rows={data.librarySummary} hiddenTitles={activeHiddenAnime} />

        <SectionHeader>Library — Cumulative</SectionHeader>
        <StackedTrendChart title="Watch Time Progress (min)" data={filteredWatchTimeProgress} />
        <StackedTrendChart title="Episodes Progress" data={filteredAnimeProgress} />
        <StackedTrendChart
          title="Cards Mined Progress"
          data={filteredCardsProgress}
          colorPalette={cardsMinedStackedColors}
        />
        <StackedTrendChart title="Words Seen Progress" data={filteredWordsProgress} />

        <SectionHeader>Patterns</SectionHeader>
        <TrendChart
          title="Watch Time by Day of Week (min)"
          data={data.patterns.watchTimeByDayOfWeek}
          color="#8aadf4"
          type="bar"
        />
        <TrendChart
          title="Watch Time by Hour (min)"
          data={data.patterns.watchTimeByHour}
          color="#c6a0f6"
          type="bar"
        />
      </div>
    </div>
  );
}
