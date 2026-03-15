import { useOverview } from '../../hooks/useOverview';
import { useStreakCalendar } from '../../hooks/useStreakCalendar';
import { HeroStats } from './HeroStats';
import { StreakCalendar } from './StreakCalendar';
import { RecentSessions } from './RecentSessions';
import { TrendChart } from '../trends/TrendChart';
import { buildOverviewSummary, buildStreakCalendar } from '../../lib/dashboard-data';
import { formatNumber } from '../../lib/formatters';

export function OverviewTab() {
  const { data, sessions, loading, error } = useOverview();
  const { calendar, loading: calLoading } = useStreakCalendar(90);

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data) return null;

  const summary = buildOverviewSummary(data);
  const streakData = buildStreakCalendar(calendar);
  const showTrackedCardNote = summary.totalTrackedCards === 0 && summary.totalSessions > 0;

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
        <h3 className="text-sm font-semibold text-ctp-text mb-3">Tracking Snapshot</h3>
        {showTrackedCardNote && (
          <div className="mb-3 rounded-lg border border-ctp-surface2 bg-ctp-surface1/50 px-3 py-2 text-xs text-ctp-subtext0">
            No tracked card-add events in the current immersion DB yet. New cards mined after this fix will show here.
          </div>
        )}
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 text-sm">
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Total Sessions</div>
            <div className="mt-1 text-xl font-semibold text-ctp-lavender">
              {formatNumber(summary.totalSessions)}
            </div>
          </div>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Episodes Today</div>
            <div className="mt-1 text-xl font-semibold text-ctp-teal">
              {formatNumber(summary.episodesToday)}
            </div>
          </div>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">All-Time Hours</div>
            <div className="mt-1 text-xl font-semibold text-ctp-mauve">
              {formatNumber(summary.allTimeHours)}
            </div>
          </div>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Active Days</div>
            <div className="mt-1 text-xl font-semibold text-ctp-peach">
              {formatNumber(summary.activeDays)}
            </div>
          </div>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">Cards Mined</div>
            <div className="mt-1 text-xl font-semibold text-ctp-green">
              {formatNumber(summary.totalTrackedCards)}
            </div>
          </div>
        </div>
      </div>

      <RecentSessions sessions={sessions} />
    </div>
  );
}
