import { useState, useEffect } from 'react';
import { useAnimeDetail } from '../../hooks/useAnimeDetail';
import { getStatsClient } from '../../hooks/useStatsApi';
import { formatDuration, formatNumber, epochDayToDate } from '../../lib/formatters';
import { StatCard } from '../layout/StatCard';
import { AnimeHeader } from './AnimeHeader';
import { EpisodeList } from './EpisodeList';
import { AnimeWordList } from './AnimeWordList';
import { AnilistSelector } from './AnilistSelector';
import { CHART_THEME } from '../../lib/chart-theme';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import type { DailyRollup } from '../../types/stats';

interface AnimeDetailViewProps {
  animeId: number;
  onBack: () => void;
  onNavigateToWord?: (wordId: number) => void;
}

type Range = 14 | 30 | 90;

function formatActiveMinutes(value: number | string) {
  const minutes = Number(value);
  return [`${Number.isFinite(minutes) ? minutes : 0} min`, 'Active Time'];
}

function AnimeWatchChart({ animeId }: { animeId: number }) {
  const [rollups, setRollups] = useState<DailyRollup[]>([]);
  const [range, setRange] = useState<Range>(30);

  useEffect(() => {
    let cancelled = false;
    getStatsClient()
      .getAnimeRollups(animeId, 90)
      .then((data) => { if (!cancelled) setRollups(data); })
      .catch(() => { if (!cancelled) setRollups([]); });
    return () => { cancelled = true; };
  }, [animeId]);

  const byDay = new Map<number, number>();
  for (const r of rollups) {
    byDay.set(r.rollupDayOrMonth, (byDay.get(r.rollupDayOrMonth) ?? 0) + r.totalActiveMin);
  }
  const chartData = Array.from(byDay.entries())
    .sort(([a], [b]) => a - b)
    .map(([day, mins]) => ({
      date: epochDayToDate(day).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }),
      minutes: Math.round(mins),
    }))
    .slice(-range);

  const ranges: Range[] = [14, 30, 90];

  if (chartData.length === 0) return null;

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-ctp-text">Watch Time</h3>
        <div className="flex gap-1">
          {ranges.map((r) => (
            <button
              key={r}
              onClick={() => setRange(r)}
              className={`px-2 py-0.5 text-xs rounded ${
                range === r
                  ? 'bg-ctp-surface2 text-ctp-text'
                  : 'text-ctp-overlay2 hover:text-ctp-subtext0'
              }`}
            >
              {r}d
            </button>
          ))}
        </div>
      </div>
      <ResponsiveContainer width="100%" height={160}>
        <BarChart data={chartData}>
          <XAxis dataKey="date" tick={{ fontSize: 10, fill: CHART_THEME.tick }} axisLine={false} tickLine={false} />
          <YAxis tick={{ fontSize: 10, fill: CHART_THEME.tick }} axisLine={false} tickLine={false} width={30} />
          <Tooltip
            contentStyle={{
              background: CHART_THEME.tooltipBg,
              border: `1px solid ${CHART_THEME.tooltipBorder}`,
              borderRadius: 6,
              color: CHART_THEME.tooltipText,
              fontSize: 12,
            }}
            labelStyle={{ color: CHART_THEME.tooltipLabel }}
            formatter={formatActiveMinutes}
          />
          <Bar dataKey="minutes" fill={CHART_THEME.barFill} radius={[3, 3, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

export function AnimeDetailView({ animeId, onBack, onNavigateToWord }: AnimeDetailViewProps) {
  const { data, loading, error, reload } = useAnimeDetail(animeId);
  const [showAnilistSelector, setShowAnilistSelector] = useState(false);

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data?.detail) return <div className="text-ctp-overlay2 p-4">Anime not found</div>;

  const { detail, episodes, anilistEntries } = data;
  const avgSessionMs = detail.totalSessions > 0
    ? Math.round(detail.totalActiveMs / detail.totalSessions)
    : 0;

  return (
    <div className="space-y-4">
      <button
        type="button"
        onClick={onBack}
        className="text-sm text-ctp-blue hover:text-ctp-sapphire transition-colors"
      >
        &larr; Back to Anime
      </button>
      <AnimeHeader
        detail={detail}
        anilistEntries={anilistEntries ?? []}
        onChangeAnilist={() => setShowAnilistSelector(true)}
      />
      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-3">
        <StatCard label="Watch Time" value={formatDuration(detail.totalActiveMs)} color="text-ctp-blue" />
        <StatCard label="Cards" value={formatNumber(detail.totalCards)} color="text-ctp-green" />
        <StatCard label="Words" value={formatNumber(detail.totalWordsSeen)} color="text-ctp-mauve" />
        <StatCard label="Sessions" value={String(detail.totalSessions)} color="text-ctp-peach" />
        <StatCard label="Avg Session" value={formatDuration(avgSessionMs)} />
      </div>
      <EpisodeList episodes={episodes} />
      <AnimeWatchChart animeId={animeId} />
      <AnimeWordList animeId={animeId} onNavigateToWord={onNavigateToWord} />
      {showAnilistSelector && (
        <AnilistSelector
          animeId={animeId}
          initialQuery={detail.canonicalTitle}
          onClose={() => setShowAnilistSelector(false)}
          onLinked={() => {
            setShowAnilistSelector(false);
            reload();
          }}
        />
      )}
    </div>
  );
}
