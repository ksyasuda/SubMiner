import { useTranslation } from '../../i18n';
import { todayLocalDay } from '../../lib/formatters';
import type { DailyRollup } from '../../types/stats';

interface QuickStatsProps {
  rollups: DailyRollup[];
}

export function QuickStats({ rollups }: QuickStatsProps) {
  const { t } = useTranslation();
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
      <h3 className="text-sm font-semibold text-ctp-text mb-3">{t('stats.quickStats.title')}</h3>
      <div className="space-y-2 text-sm">
        <div className="flex justify-between">
          <span className="text-ctp-subtext0">{t('stats.quickStats.streak')}</span>
          <span className="text-ctp-peach font-medium">
            {streak} day{streak !== 1 ? 's' : ''}
          </span>
        </div>
        <div className="flex justify-between">
          <span className="text-ctp-subtext0">{t('stats.quickStats.avgPerDay')}</span>
          <span className="text-ctp-text">{avgMinPerDay}m</span>
        </div>
        <div className="flex justify-between">
          <span className="text-ctp-subtext0">{t('stats.quickStats.cardsThisWeek')}</span>
          <span className="text-ctp-cards-mined font-medium">{weekCards}</span>
        </div>
      </div>
    </div>
  );
}
