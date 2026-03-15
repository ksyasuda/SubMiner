import { todayLocalDay } from '../../lib/formatters';
import type { DailyRollup } from '../../types/stats';

interface QuickStatsProps {
  rollups: DailyRollup[];
}

export function QuickStats({ rollups }: QuickStatsProps) {
  const daysWithActivity = new Set(
    rollups.filter((r) => r.totalActiveMin > 0).map((r) => r.rollupDayOrMonth),
  );
  const today = todayLocalDay();
  const streakStart = daysWithActivity.has(today) ? today : today - 1;
  let streak = 0;
  for (let d = streakStart; daysWithActivity.has(d); d--) {
    streak++;
  }

  const weekStart = today - 6;
  const weekRollups = rollups.filter((r) => r.rollupDayOrMonth >= weekStart);
  const weekMinutes = weekRollups.reduce((sum, r) => sum + r.totalActiveMin, 0);
  const weekCards = weekRollups.reduce((sum, r) => sum + r.totalCards, 0);
  const avgMinPerDay = Math.round(weekMinutes / 7);

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-sm font-semibold text-ctp-text mb-3">Quick Stats</h3>
      <div className="space-y-2 text-sm">
        <div className="flex justify-between">
          <span className="text-ctp-subtext0">Streak</span>
          <span className="text-ctp-peach font-medium">
            {streak} day{streak !== 1 ? 's' : ''}
          </span>
        </div>
        <div className="flex justify-between">
          <span className="text-ctp-subtext0">Avg/day this week</span>
          <span className="text-ctp-text">{avgMinPerDay}m</span>
        </div>
        <div className="flex justify-between">
          <span className="text-ctp-subtext0">Cards this week</span>
          <span className="text-ctp-green font-medium">{weekCards}</span>
        </div>
      </div>
    </div>
  );
}
