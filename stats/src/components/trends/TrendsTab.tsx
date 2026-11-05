import { useMemo, useState, type ReactNode } from 'react';
import { useTrends, type GroupBy, type TimeRange } from '../../hooks/useTrends';
import { formatDuration, formatNumber } from '../../lib/formatters';
import type { SessionsDayFilter } from '../../lib/stats-navigation';
import type { LibrarySummaryRow, TrendBucket, TrendChartPoint } from '../../types/stats';
import { DateRangeSelector } from './DateRangeSelector';
import { DayHourHeatmap } from './DayHourHeatmap';
import { FrequencyBandChart } from './FrequencyBandChart';
import { LibrarySummarySection } from './LibrarySummarySection';
import { SummaryStrip } from './SummaryStrip';
import { TitleFilter } from './TitleFilter';
import { TitleTimeline } from './TitleTimeline';
import { startFromZero, TrendChart } from './TrendChart';
import {
  RUNNING_TOTAL_METRICS,
  loadGroupBy,
  loadRange,
  loadRunningTotalMetric,
  loadShowEmptyDays,
  saveGroupBy,
  saveRange,
  saveRunningTotalMetric,
  saveShowEmptyDays,
  type RunningTotalMetric,
} from './trends-preferences';

interface TrendsTabProps {
  onNavigateToAnime?: (animeId: number) => void;
  onNavigateToMediaDetail?: (videoId: number) => void;
  onOpenSessionsForDays?: (filter: SessionsDayFilter) => void;
}

const WATCH_TIME_COLOR = '#8aadf4';
const WORDS_COLOR = '#8bd5ca';
const NEW_WORDS_COLOR = '#c6a0f6';
const EPISODES_COLOR = '#91d7e3';
const LOOKUPS_COLOR = '#f5a97f';
const READING_SPEED_COLOR = '#a6da95';
const CARDS_COLOR = 'var(--color-ctp-cards-mined)';

const RUNNING_TOTAL_LABELS: Record<RunningTotalMetric, string> = {
  watchTime: 'Watch time',
  words: 'Words seen',
  newWords: 'New words',
  cards: 'Cards mined',
  episodes: 'Episodes finished',
  lookups: 'Lookups',
};

const RUNNING_TOTAL_COLORS: Record<RunningTotalMetric, string> = {
  watchTime: WATCH_TIME_COLOR,
  words: WORDS_COLOR,
  newWords: NEW_WORDS_COLOR,
  cards: CARDS_COLOR,
  episodes: EPISODES_COLOR,
  lookups: LOOKUPS_COLOR,
};

function SectionHeader({ children, hint }: { children: ReactNode; hint?: string }) {
  return (
    <div className="col-span-full mt-6 mb-2 flex items-center gap-3">
      <h3 className="text-ctp-subtext0 text-xs font-semibold uppercase tracking-widest shrink-0">
        {children}
      </h3>
      <div className="flex-1 h-px bg-gradient-to-r from-ctp-surface1 to-transparent" />
      {hint ? <span className="shrink-0 text-[11px] text-ctp-overlay2">{hint}</span> : null}
    </div>
  );
}

/** Chart values in hours, so axis ticks land on round hours instead of odd minute counts. */
function toHours(points: TrendChartPoint[]): TrendChartPoint[] {
  return points.map((point) => ({ ...point, value: point.value / 60 }));
}

const formatHourTick = (hours: number): string => `${+hours.toFixed(2)}h`;
const formatHourValue = (hours: number): string => formatDuration(hours * 3_600_000);

function bucketSessionsLabel(bucket: TrendBucket, groupBy: GroupBy): string {
  return groupBy === 'week' ? `the week of ${bucket.label}` : bucket.label;
}

export function TrendsTab({
  onNavigateToAnime,
  onNavigateToMediaDetail,
  onOpenSessionsForDays,
}: TrendsTabProps = {}) {
  const [range, setRange] = useState<TimeRange>(() => loadRange());
  const [groupBy, setGroupBy] = useState<GroupBy>(() => loadGroupBy());
  const [showEmptyDays, setShowEmptyDays] = useState(() => loadShowEmptyDays());
  const [runningTotalMetric, setRunningTotalMetric] = useState(() => loadRunningTotalMetric());
  const [selectedTitle, setSelectedTitle] = useState<string | null>(null);
  const [titleQuery, setTitleQuery] = useState('');
  const { data, loading, refreshing, error } = useTrends(
    range,
    groupBy,
    showEmptyDays,
    selectedTitle,
  );

  const updateRange = (next: TimeRange) => {
    setRange(next);
    saveRange(next);
  };
  const updateGroupBy = (next: GroupBy) => {
    setGroupBy(next);
    saveGroupBy(next);
  };
  const updateShowEmptyDays = (show: boolean) => {
    setShowEmptyDays(show);
    saveShowEmptyDays(show);
  };
  const updateRunningTotalMetric = (metric: RunningTotalMetric) => {
    setRunningTotalMetric(metric);
    saveRunningTotalMetric(metric);
  };
  const toggleSelectedTitle = (title: string) => {
    setSelectedTitle((current) => (current === title ? null : title));
  };

  const openTitle = (row: LibrarySummaryRow) => {
    if (row.animeId !== null) {
      onNavigateToAnime?.(row.animeId);
    } else if (row.videoId !== null) {
      onNavigateToMediaDetail?.(row.videoId);
    }
  };

  const allTitles = useMemo(() => data?.librarySummary.map((row) => row.title) ?? [], [data]);
  const visibleTitleRows = useMemo(() => {
    const rows = data?.librarySummary ?? [];
    const query = titleQuery.trim().toLowerCase();
    return query ? rows.filter((row) => row.title.toLowerCase().includes(query)) : rows;
  }, [data, titleQuery]);

  const controls = (
    <div className="flex flex-wrap items-center gap-x-4 gap-y-2">
      <DateRangeSelector
        range={range}
        groupBy={groupBy}
        fillEmpty={showEmptyDays}
        shownDays={data?.window ?? null}
        onRangeChange={updateRange}
        onGroupByChange={updateGroupBy}
        onFillEmptyChange={updateShowEmptyDays}
      />
      <TitleFilter
        titles={allTitles}
        query={titleQuery}
        selectedTitle={selectedTitle}
        onQueryChange={setTitleQuery}
        onSelectTitle={setSelectedTitle}
      />
      {refreshing && data ? (
        <span role="status" className="text-[11px] text-ctp-overlay2">
          Updating...
        </span>
      ) : null}
    </div>
  );

  if (loading) {
    return (
      <div className="space-y-4">
        {controls}
        <div className="text-ctp-overlay2 p-4">Loading...</div>
      </div>
    );
  }
  if (!data) {
    return (
      <div className="space-y-4">
        {controls}
        {error ? <div className="text-ctp-red p-4">Error: {error}</div> : null}
      </div>
    );
  }

  const openBucketSessions = (_label: string, index: number) => {
    const bucket = data.buckets[index];
    if (!bucket) return;
    onOpenSessionsForDays?.({
      startDay: bucket.startDay,
      endDay: bucket.endDay,
      label: bucketSessionsLabel(bucket, groupBy),
    });
  };
  const barClickProps = onOpenSessionsForDays ? { onBarClick: openBucketSessions } : {};

  const windowDays = data.window.endDay - data.window.startDay + 1;
  const runningTotalOptions = RUNNING_TOTAL_METRICS.filter(
    (metric) => metric !== 'newWords' || data.learning !== null,
  );
  const activeRunningTotal = runningTotalOptions.includes(runningTotalMetric)
    ? runningTotalMetric
    : 'watchTime';
  const isWatchTimeTotal = activeRunningTotal === 'watchTime';

  return (
    <div className="space-y-4" aria-busy={refreshing}>
      {controls}
      {error ? <div className="text-sm text-ctp-red">Error: {error}</div> : null}
      <SummaryStrip summary={data.summary} days={windowDays} />
      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-6 gap-4">
        <SectionHeader
          hint={onOpenSessionsForDays ? 'Click a bar to open its sessions' : undefined}
        >
          Activity per {groupBy}
        </SectionHeader>
        <div className="xl:col-span-3">
          <TrendChart
            title="Watch Time (hours)"
            data={toHours(data.activity.watchTime)}
            color={WATCH_TIME_COLOR}
            type="bar"
            tickFormatter={formatHourTick}
            valueFormatter={formatHourValue}
            {...barClickProps}
          />
        </div>
        <div className="xl:col-span-3">
          <TrendChart
            title="Cards Mined"
            data={data.activity.cards}
            color={CARDS_COLOR}
            type="bar"
            {...barClickProps}
          />
        </div>
        <div className="xl:col-span-3">
          <TrendChart
            title="Words Seen"
            data={data.activity.words}
            color={WORDS_COLOR}
            type="bar"
            {...barClickProps}
          />
        </div>
        <div className="xl:col-span-3">
          <TrendChart
            title="Episodes Finished"
            data={data.activity.episodes}
            color={EPISODES_COLOR}
            type="bar"
            {...barClickProps}
          />
        </div>

        {data.learning ? (
          <>
            <SectionHeader hint="Dictionary forms, names excluded">Learning</SectionHeader>
            <div className="xl:col-span-3">
              <TrendChart
                title={
                  groupBy === 'day'
                    ? 'New Words per 1,000 Seen (rolling 7 days)'
                    : 'New Words per 1,000 Seen'
                }
                data={data.learning.newWordsPerThousand}
                color={NEW_WORDS_COLOR}
                type="line"
              />
            </div>
            <div className="xl:col-span-3">
              <FrequencyBandChart data={data.learning.frequencyBands} />
            </div>
          </>
        ) : (
          <div className="col-span-full text-xs text-ctp-overlay2">
            New words are tracked across your whole library, so the Learning charts are hidden while
            filtering to one title.
          </div>
        )}

        <SectionHeader
          hint={
            groupBy === 'day'
              ? 'Rolling 7 days, gaps mean too little data'
              : 'Gaps mean too little data'
          }
        >
          Efficiency
        </SectionHeader>
        <div className="xl:col-span-2">
          <TrendChart
            title="Words / Min"
            data={data.ratios.readingSpeed}
            color={READING_SPEED_COLOR}
            type="line"
          />
        </div>
        <div className="xl:col-span-2">
          <TrendChart
            title="Cards / Hour"
            data={data.ratios.cardsPerHour}
            color={CARDS_COLOR}
            type="line"
          />
        </div>
        <div className="xl:col-span-2">
          <TrendChart
            title="Lookups / 100 Words"
            data={data.ratios.lookupsPerHundred}
            color={LOOKUPS_COLOR}
            type="line"
          />
        </div>

        <SectionHeader>Running Total</SectionHeader>
        <div className="col-span-full">
          <div className="mb-2 flex flex-wrap gap-1" role="group" aria-label="Running total metric">
            {runningTotalOptions.map((metric) => (
              <button
                key={metric}
                type="button"
                aria-pressed={metric === activeRunningTotal}
                className={`rounded-md border px-2 py-1 text-[11px] transition ${
                  metric === activeRunningTotal
                    ? 'border-ctp-surface2 bg-ctp-surface2 text-ctp-text'
                    : 'border-ctp-surface1 text-ctp-overlay2 hover:text-ctp-subtext0'
                }`}
                onClick={() => updateRunningTotalMetric(metric)}
              >
                {RUNNING_TOTAL_LABELS[metric]}
              </button>
            ))}
          </div>
          <TrendChart
            title={`${RUNNING_TOTAL_LABELS[activeRunningTotal]}, total since the start of this range`}
            data={startFromZero(
              isWatchTimeTotal
                ? toHours(data.progress.watchTime)
                : data.progress[activeRunningTotal],
            )}
            color={RUNNING_TOTAL_COLORS[activeRunningTotal]}
            type="line"
            {...(isWatchTimeTotal
              ? { tickFormatter: formatHourTick, valueFormatter: formatHourValue }
              : {})}
          />
        </div>

        <SectionHeader>Patterns</SectionHeader>
        <DayHourHeatmap grid={data.watchTimeByDayHour} />

        <SectionHeader
          hint={`${formatNumber(data.librarySummary.length)} titles. Click a title to filter the page, ↗ to open it`}
        >
          Titles
        </SectionHeader>
        <TitleTimeline
          rows={visibleTitleRows}
          points={data.titleTimeline}
          window={data.window}
          selectedTitle={selectedTitle}
          onSelectTitle={toggleSelectedTitle}
          onOpenTitle={openTitle}
        />
        <LibrarySummarySection
          rows={visibleTitleRows}
          selectedTitle={selectedTitle}
          onSelectTitle={toggleSelectedTitle}
          onOpenTitle={openTitle}
        />
      </div>
    </div>
  );
}
