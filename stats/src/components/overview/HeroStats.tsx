import { useTranslation } from '../../i18n';
import { StatCard } from '../layout/StatCard';
import { formatDuration, formatNumber, todayLocalDay, localDayFromMs } from '../../lib/formatters';
import type { OverviewSummary } from '../../lib/dashboard-data';
import type { SessionSummary } from '../../types/stats';

interface HeroStatsProps {
  summary: OverviewSummary;
  sessions: SessionSummary[];
}

export function HeroStats({ summary, sessions }: HeroStatsProps) {
  const { t } = useTranslation();
  const today = todayLocalDay();
  const sessionsToday = sessions.filter((s) => localDayFromMs(s.startedAtMs) === today).length;

  return (
    <div className="grid grid-cols-2 xl:grid-cols-6 gap-3">
      <StatCard
        label={t('stats.hero.watchTimeToday')}
        value={formatDuration(summary.todayActiveMs)}
        color="text-ctp-blue"
      />
      <StatCard
        label={t('stats.hero.cardsMinedToday')}
        value={formatNumber(summary.todayCards)}
        color="text-ctp-cards-mined"
      />
      <StatCard
        label={t('stats.hero.sessionsToday')}
        value={formatNumber(sessionsToday)}
        color="text-ctp-lavender"
      />
      <StatCard
        label={t('stats.hero.episodesToday')}
        value={formatNumber(summary.episodesToday)}
        color="text-ctp-teal"
      />
      <StatCard label={t('stats.hero.currentStreak')} value={`${summary.streakDays}d`} color="text-ctp-peach" />
      <StatCard
        label={t('stats.hero.activeTitles')}
        value={formatNumber(summary.activeAnimeCount)}
        color="text-ctp-mauve"
      />
    </div>
  );
}
