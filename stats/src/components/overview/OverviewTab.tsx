import { useState, useEffect } from 'react';
import { useOverview } from '../../hooks/useOverview';
import { useStreakCalendar } from '../../hooks/useStreakCalendar';
import { HeroStats } from './HeroStats';
import { StreakCalendar } from './StreakCalendar';
import { RecentSessions } from './RecentSessions';
import { TrackingSnapshot } from './TrackingSnapshot';
import { TrendChart } from '../trends/TrendChart';
import { buildOverviewSummary, buildStreakCalendar } from '../../lib/dashboard-data';
import { apiClient } from '../../lib/api-client';
import { getStatsClient } from '../../hooks/useStatsApi';
import {
  confirmSessionDelete,
  confirmDayGroupDelete,
  confirmAnimeGroupDelete,
} from '../../lib/delete-confirm';
import type { SessionSummary } from '../../types/stats';

interface OverviewTabProps {
  onNavigateToMediaDetail: (videoId: number, sessionId?: number | null) => void;
  onNavigateToSession: (sessionId: number) => void;
  isActive?: boolean;
}

export function OverviewTab({
  onNavigateToMediaDetail,
  onNavigateToSession,
  isActive = true,
}: OverviewTabProps) {
  const [refreshKey, setRefreshKey] = useState(0);
  const { data, sessions, setSessions, loading, error } = useOverview(refreshKey);
  const { calendar, loading: calLoading } = useStreakCalendar(90, refreshKey);
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
  }, [refreshKey]);

  const handleDeleteSession = async (session: SessionSummary) => {
    if (!(await confirmSessionDelete())) return;
    setDeleteError(null);
    setDeletingIds((prev) => new Set(prev).add(session.sessionId));
    try {
      await apiClient.deleteSession(session.sessionId);
      setSessions((prev) => prev.filter((s) => s.sessionId !== session.sessionId));
      setRefreshKey((prev) => prev + 1);
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
    if (!(await confirmDayGroupDelete(dayLabel, daySessions.length))) return;
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
      setRefreshKey((prev) => prev + 1);
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
    if (!(await confirmAnimeGroupDelete(title, groupSessions.length))) return;
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
      setRefreshKey((prev) => prev + 1);
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

  const errorNotice = error ? (
    <div
      role="alert"
      className="flex items-center justify-between gap-3 rounded-lg border border-ctp-red/30 bg-ctp-red/10 p-3 text-sm text-ctp-red"
    >
      <span>
        {data
          ? 'Could not refresh Overview. Displayed totals may be out of date.'
          : 'Could not load Overview.'}{' '}
        {error}
      </span>
      <button
        type="button"
        onClick={() => setRefreshKey((prev) => prev + 1)}
        className="shrink-0 rounded border border-ctp-red/30 px-3 py-1 hover:bg-ctp-red/10"
      >
        Retry
      </button>
    </div>
  ) : null;

  if (loading && !data) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (!data) return errorNotice;

  const summary = buildOverviewSummary(data);
  const streakData = buildStreakCalendar(calendar);
  const showTrackedCardNote = summary.totalTrackedCards === 0 && summary.activeDays > 0;

  return (
    <div className="space-y-4">
      {errorNotice}
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

      <TrackingSnapshot
        summary={summary}
        showTrackedCardNote={showTrackedCardNote}
        knownWordsSummary={knownWordsSummary}
      />

      {deleteError ? <div className="text-sm text-ctp-red">{deleteError}</div> : null}

      <RecentSessions
        sessions={sessions}
        onNavigateToMediaDetail={onNavigateToMediaDetail}
        onNavigateToSession={onNavigateToSession}
        onDeleteSession={handleDeleteSession}
        onDeleteDayGroup={handleDeleteDayGroup}
        onDeleteAnimeGroup={handleDeleteAnimeGroup}
        deletingIds={deletingIds}
        isActive={isActive}
      />
    </div>
  );
}
