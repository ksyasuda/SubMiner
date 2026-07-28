import { useState, useEffect, useRef } from 'react';
import { useAnimeDetail } from '../../hooks/useAnimeDetail';
import { getStatsClient } from '../../hooks/useStatsApi';
import { confirmAnimeDelete } from '../../lib/delete-confirm';
import { epochDayToDate } from '../../lib/formatters';
import { AnimeHeader } from './AnimeHeader';
import { EpisodeList } from './EpisodeList';
import { AnimeWordList } from './AnimeWordList';
import { AnilistSelector } from './AnilistSelector';
import { AnimeOverviewStats } from './AnimeOverviewStats';
import { CHART_THEME } from '../../lib/chart-theme';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import type { DailyRollup } from '../../types/stats';

interface AnimeDetailViewProps {
  animeId: number;
  onBack: () => void;
  onNavigateToWord?: (wordId: number) => void;
  onOpenEpisodeDetail?: (videoId: number) => void;
  /** Called after the whole library entry is deleted, so the caller can refresh. */
  onAnimeDeleted?: () => void;
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
      .then((data) => {
        if (!cancelled) setRollups(data);
      })
      .catch(() => {
        if (!cancelled) setRollups([]);
      });
    return () => {
      cancelled = true;
    };
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
          <XAxis
            dataKey="date"
            tick={{ fontSize: 10, fill: CHART_THEME.tick }}
            axisLine={false}
            tickLine={false}
          />
          <YAxis
            tick={{ fontSize: 10, fill: CHART_THEME.tick }}
            axisLine={false}
            tickLine={false}
            width={30}
          />
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

function useAnimeKnownWords(animeId: number) {
  const [summary, setSummary] = useState<{
    totalUniqueWords: number;
    knownWordCount: number;
  } | null>(null);
  useEffect(() => {
    let cancelled = false;
    getStatsClient()
      .getAnimeKnownWordsSummary(animeId)
      .then((data) => {
        if (!cancelled) setSummary(data);
      })
      .catch(() => {
        if (!cancelled) setSummary(null);
      });
    return () => {
      cancelled = true;
    };
  }, [animeId]);
  return summary;
}

export function AnimeDetailView({
  animeId,
  onBack,
  onNavigateToWord,
  onOpenEpisodeDetail,
  onAnimeDeleted,
}: AnimeDetailViewProps) {
  const { data, loading, error, reload } = useAnimeDetail(animeId);
  const [showAnilistSelector, setShowAnilistSelector] = useState(false);
  const [coverRetryToken, setCoverRetryToken] = useState(0);
  const [isDeletingAnime, setIsDeletingAnime] = useState(false);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const isDeletingAnimeRef = useRef(false);
  const knownWordsSummary = useAnimeKnownWords(animeId);

  useEffect(() => {
    setCoverRetryToken(0);
  }, [animeId]);

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data?.detail) return <div className="text-ctp-overlay2 p-4">Anime not found</div>;

  const { detail, episodes, anilistEntries } = data;

  const handleDeleteAnime = async () => {
    if (isDeletingAnimeRef.current) return;
    isDeletingAnimeRef.current = true;
    // Cleared up front so cancelling a retry doesn't leave the previous
    // attempt's failure on screen.
    setDeleteError(null);
    let confirmed = false;
    try {
      confirmed = await confirmAnimeDelete(detail.canonicalTitle, detail.episodeCount);
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to confirm delete.');
      isDeletingAnimeRef.current = false;
      return;
    }
    if (!confirmed) {
      isDeletingAnimeRef.current = false;
      return;
    }

    setDeleteError(null);
    setIsDeletingAnime(true);
    try {
      await getStatsClient().deleteAnime(animeId);
      onAnimeDeleted?.();
      onBack();
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete this title.');
      setIsDeletingAnime(false);
    } finally {
      isDeletingAnimeRef.current = false;
    }
  };

  return (
    <div className="space-y-4">
      <button
        type="button"
        onClick={onBack}
        className="text-sm text-ctp-blue hover:text-ctp-sapphire transition-colors"
      >
        &larr; Back to Library
      </button>
      <AnimeHeader
        detail={detail}
        anilistEntries={anilistEntries ?? []}
        coverRetryToken={coverRetryToken}
        onChangeAnilist={() => setShowAnilistSelector(true)}
        onDeleteAnime={() => void handleDeleteAnime()}
        isDeletingAnime={isDeletingAnime}
      />
      {deleteError ? <div className="text-sm text-ctp-red">{deleteError}</div> : null}
      <AnimeOverviewStats detail={detail} knownWordsSummary={knownWordsSummary} />
      <EpisodeList
        episodes={episodes}
        onOpenDetail={onOpenEpisodeDetail ? (videoId) => onOpenEpisodeDetail(videoId) : undefined}
      />
      <AnimeWatchChart animeId={animeId} />
      <AnimeWordList animeId={animeId} onNavigateToWord={onNavigateToWord} />
      {showAnilistSelector && (
        <AnilistSelector
          animeId={animeId}
          initialQuery={detail.canonicalTitle}
          onClose={() => setShowAnilistSelector(false)}
          onLinked={() => {
            setShowAnilistSelector(false);
            setCoverRetryToken((value) => value + 1);
            reload();
          }}
        />
      )}
    </div>
  );
}
