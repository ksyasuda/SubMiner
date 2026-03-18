import { useState, useEffect } from 'react';
import { useOverview } from '../../hooks/useOverview';
import { useStreakCalendar } from '../../hooks/useStreakCalendar';
import { HeroStats } from './HeroStats';
import { StreakCalendar } from './StreakCalendar';
import { RecentSessions } from './RecentSessions';
import { TrendChart } from '../trends/TrendChart';
import { buildOverviewSummary, buildStreakCalendar } from '../../lib/dashboard-data';
import { formatNumber } from '../../lib/formatters';
import { apiClient } from '../../lib/api-client';
import { getStatsClient } from '../../hooks/useStatsApi';
import { Tooltip } from '../layout/Tooltip';
import {
  confirmSessionDelete,
  confirmDayGroupDelete,
  confirmAnimeGroupDelete,
} from '../../lib/delete-confirm';
import type { SessionSummary } from '../../types/stats';

interface OverviewTabProps {
  onNavigateToMediaDetail: (videoId: number, sessionId?: number | null) => void;
  onNavigateToSession: (sessionId: number) => void;
}

export function OverviewTab({ onNavigateToMediaDetail, onNavigateToSession }: OverviewTabProps) {
  const { data, sessions, setSessions, loading, error } = useOverview();
  const { calendar, loading: calLoading } = useStreakCalendar(90);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [deletingIds, setDeletingIds] = useState<Set<number>>(new Set());
  const [knownWordsSummary, setKnownWordsSummary] = useState<{
    totalUniqueWords: number;
    knownWordCount: number;
  } | null>(null);

  useEffect(() => {
    let cancelled = false;
    getStatsClient()
      .getKnownWordsSummary()
      .then((data) => {
        if (!cancelled) setKnownWordsSummary(data);
      })
      .catch(() => {
        if (!cancelled) setKnownWordsSummary(null);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const handleDeleteSession = async (session: SessionSummary) => {
    if (!confirmSessionDelete()) return;
    setDeleteError(null);
    setDeletingIds((prev) => new Set(prev).add(session.sessionId));
    try {
      await apiClient.deleteSession(session.sessionId);
      setSessions((prev) => prev.filter((s) => s.sessionId !== session.sessionId));
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete session.');
    } finally {
      setDeletingIds((prev) => {
        const next = new Set(prev);
        next.delete(session.sessionId);
        return next;
      });
    }
  };

  const handleDeleteDayGroup = async (dayLabel: string, daySessions: SessionSummary[]) => {
    if (!confirmDayGroupDelete(dayLabel, daySessions.length)) return;
    setDeleteError(null);
    const ids = daySessions.map((s) => s.sessionId);
    setDeletingIds((prev) => {
      const next = new Set(prev);
      for (const id of ids) next.add(id);
      return next;
    });
    try {
      await apiClient.deleteSessions(ids);
      const idSet = new Set(ids);
      setSessions((prev) => prev.filter((s) => !idSet.has(s.sessionId)));
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete sessions.');
    } finally {
      setDeletingIds((prev) => {
        const next = new Set(prev);
        for (const id of ids) next.delete(id);
        return next;
      });
    }
  };

  const handleDeleteAnimeGroup = async (groupSessions: SessionSummary[]) => {
    const title =
      groupSessions[0]?.animeTitle ?? groupSessions[0]?.canonicalTitle ?? 'Unknown Media';
    if (!confirmAnimeGroupDelete(title, groupSessions.length)) return;
    setDeleteError(null);
    const ids = groupSessions.map((s) => s.sessionId);
    setDeletingIds((prev) => {
      const next = new Set(prev);
      for (const id of ids) next.add(id);
      return next;
    });
    try {
      await apiClient.deleteSessions(ids);
      const idSet = new Set(ids);
      setSessions((prev) => prev.filter((s) => !idSet.has(s.sessionId)));
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete sessions.');
    } finally {
      setDeletingIds((prev) => {
        const next = new Set(prev);
        for (const id of ids) next.delete(id);
        return next;
      });
    }
  };

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data) return null;

  const summary = buildOverviewSummary(data);
  const streakData = buildStreakCalendar(calendar);
  const showTrackedCardNote = summary.totalTrackedCards === 0 && summary.activeDays > 0;

  return (
    <div className="space-y-4">
      <HeroStats summary={summary} sessions={sessions} />

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <TrendChart
          title="Last 14 Days Watch Time (min)"
          data={summary.recentWatchTime}
          color="#8aadf4"
          type="bar"
        />
        {!calLoading && <StreakCalendar data={streakData} />}
      </div>

      <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
        <h3 className="text-sm font-semibold text-ctp-text">Tracking Snapshot</h3>
        <p className="mt-1 mb-3 text-xs text-ctp-overlay2">
          Lifetime totals sourced from summary tables.
        </p>
        {showTrackedCardNote && (
          <div className="mb-3 rounded-lg border border-ctp-surface2 bg-ctp-surface1/50 px-3 py-2 text-xs text-ctp-subtext0">
            No lifetime card totals in the summary table yet. New cards mined after this fix will
            appear here.
          </div>
        )}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-sm">
          <Tooltip text="Total immersion sessions recorded across all time">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Sessions</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-lavender">
                {formatNumber(summary.totalSessions)}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Total active watch time across all sessions">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Watch Time</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-mauve">
                {summary.allTimeMinutes < 60
                  ? `${summary.allTimeMinutes}m`
                  : `${(summary.allTimeMinutes / 60).toFixed(1)}h`}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Number of distinct days with at least one session">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Active Days</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-peach">
                {formatNumber(summary.activeDays)}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Average active watch time per session in minutes">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Avg Session</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-yellow">
                {formatNumber(summary.averageSessionMinutes)}
                <span className="text-sm text-ctp-overlay2 ml-0.5">min</span>
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Total unique episodes (videos) watched across all anime">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Episodes</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-blue">
                {formatNumber(summary.totalEpisodesWatched)}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Number of anime series fully completed">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Anime</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-sapphire">
                {formatNumber(summary.totalAnimeCompleted)}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Total Anki cards mined from subtitle lines across all sessions">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Cards Mined</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-green">
                {formatNumber(summary.totalTrackedCards)}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Percentage of dictionary lookups that matched a known word">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Lookup Rate</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-flamingo">
                {summary.lookupRate != null ? `${summary.lookupRate}%` : '—'}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Total word occurrences encountered in today's sessions">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Words Today</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-sky">
                {formatNumber(summary.todayWords)}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Unique words seen for the first time today">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">
                New Words Today
              </div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-rosewater">
                {formatNumber(summary.newWordsToday)}
              </div>
            </div>
          </Tooltip>
          <Tooltip text="Unique words seen for the first time this week">
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">New Words</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-pink">
                {formatNumber(summary.newWordsThisWeek)}
              </div>
            </div>
          </Tooltip>
          {knownWordsSummary && knownWordsSummary.totalUniqueWords > 0 && (
            <>
              <Tooltip text="Words matched against your known-words list out of all unique words seen">
                <div className="rounded-lg bg-ctp-surface1/60 p-3">
                  <div className="text-xs uppercase tracking-wide text-ctp-overlay2">
                    Known Words
                  </div>
                  <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-green">
                    {formatNumber(knownWordsSummary.knownWordCount)}
                    <span className="text-sm text-ctp-overlay2 ml-1">
                      / {formatNumber(knownWordsSummary.totalUniqueWords)}
                    </span>
                  </div>
                </div>
              </Tooltip>
            </>
          )}
        </div>
      </div>

      {deleteError ? <div className="text-sm text-ctp-red">{deleteError}</div> : null}

      <RecentSessions
        sessions={sessions}
        onNavigateToMediaDetail={onNavigateToMediaDetail}
        onNavigateToSession={onNavigateToSession}
        onDeleteSession={handleDeleteSession}
        onDeleteDayGroup={handleDeleteDayGroup}
        onDeleteAnimeGroup={handleDeleteAnimeGroup}
        deletingIds={deletingIds}
      />
    </div>
  );
}
