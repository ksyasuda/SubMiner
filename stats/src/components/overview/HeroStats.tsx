import { StatCard } from '../layout/StatCard';
import { formatDuration, formatNumber, todayLocalDay, localDayFromMs } from '../../lib/formatters';
import type { OverviewSummary } from '../../lib/dashboard-data';
import type { SessionSummary } from '../../types/stats';

interface HeroStatsProps {
  summary: OverviewSummary;
  sessions: SessionSummary[];
}

export function HeroStats({ summary, sessions }: HeroStatsProps) {
  const today = todayLocalDay();
  const sessionsToday = sessions.filter((s) => localDayFromMs(s.startedAtMs) === today).length;

  return (
    <div className="grid grid-cols-2 xl:grid-cols-6 gap-3">
      <StatCard
        label="Watch Time Today"
        value={formatDuration(summary.todayActiveMs)}
        color="text-ctp-blue"
      />
      <StatCard
        label="Cards Mined Today"
        value={formatNumber(summary.todayCards)}
        color="text-ctp-cards-mined"
      />
      <StatCard
        label="Sessions Today"
        value={formatNumber(sessionsToday)}
        color="text-ctp-lavender"
      />
      <StatCard
        label="Episodes Today"
        value={formatNumber(summary.episodesToday)}
        color="text-ctp-teal"
      />
      <StatCard label="Current Streak" value={`${summary.streakDays}d`} color="text-ctp-peach" />
      <StatCard
        label="Active Titles"
        value={formatNumber(summary.activeAnimeCount)}
        color="text-ctp-mauve"
      />
    </div>
  );
}
