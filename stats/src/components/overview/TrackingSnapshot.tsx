import { useTranslation } from '../../i18n';
import type { OverviewSummary } from '../../lib/dashboard-data';
import { formatNumber } from '../../lib/formatters';
import { Tooltip } from '../layout/Tooltip';

interface KnownWordsSummary {
  totalUniqueWords: number;
  knownWordCount: number;
}

interface TrackingSnapshotProps {
  summary: OverviewSummary;
  showTrackedCardNote?: boolean;
  knownWordsSummary: KnownWordsSummary | null;
}

export function TrackingSnapshot({
  summary,
  showTrackedCardNote = false,
  knownWordsSummary,
}: TrackingSnapshotProps) {
  const { t } = useTranslation();
  const knownWordPercent =
    knownWordsSummary && knownWordsSummary.totalUniqueWords > 0
      ? Math.round((knownWordsSummary.knownWordCount / knownWordsSummary.totalUniqueWords) * 100)
      : null;

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-sm font-semibold text-ctp-text">{t('stats.tracking.title')}</h3>
      <p className="mt-1 mb-3 text-xs text-ctp-overlay2">
        {t('stats.tracking.subtitle')}
      </p>
      {showTrackedCardNote && (
        <div className="mb-3 rounded-lg border border-ctp-surface2 bg-ctp-surface1/50 px-3 py-2 text-xs text-ctp-subtext0">
          No lifetime card totals in the summary table yet. New cards mined after this fix will
          appear here.
        </div>
      )}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-sm">
        <Tooltip text={t('stats.tracking.sessionsTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.sessions')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-lavender">
              {formatNumber(summary.totalSessions)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.watchTimeTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.watchTime')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-mauve">
              {summary.allTimeMinutes < 60
                ? `${summary.allTimeMinutes}m`
                : `${(summary.allTimeMinutes / 60).toFixed(1)}h`}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.activeDaysTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.activeDays')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-peach">
              {formatNumber(summary.activeDays)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.avgSessionTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.avgSession')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-yellow">
              {formatNumber(summary.averageSessionMinutes)}
              <span className="text-sm text-ctp-overlay2 ml-0.5">min</span>
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.episodesTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.episodes')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-blue">
              {formatNumber(summary.totalEpisodesWatched)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.titlesTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.titles')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-sapphire">
              {formatNumber(summary.totalAnimeCompleted)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.cardsMinedTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.cardsMined')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-cards-mined">
              {formatNumber(summary.totalTrackedCards)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.lookupRateTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.lookupRate')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-flamingo">
              {summary.lookupRate?.shortValue ?? '—'}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.wordsTodayTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.wordsToday')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-sky">
              {formatNumber(summary.todayTokens)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.newWordsTodayTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.newWordsToday')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-rosewater">
              {formatNumber(summary.newWordsToday)}
            </div>
          </div>
        </Tooltip>
        <Tooltip text={t('stats.tracking.newWordsTooltip')}>
          <div className="rounded-lg bg-ctp-surface1/60 p-3">
            <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.newWords')}</div>
            <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-pink">
              {formatNumber(summary.newWordsThisWeek)}
            </div>
          </div>
        </Tooltip>
        {knownWordsSummary && knownWordsSummary.totalUniqueWords > 0 && (
          <Tooltip text={t('stats.tracking.knownWordsTooltip')}>
            <div className="rounded-lg bg-ctp-surface1/60 p-3">
              <div className="text-xs uppercase tracking-wide text-ctp-overlay2">{t('stats.tracking.knownWords')}</div>
              <div className="mt-1 text-xl font-semibold font-mono tabular-nums text-ctp-green">
                {formatNumber(knownWordsSummary.knownWordCount)}
                <span className="text-sm text-ctp-overlay2 ml-1">
                  / {formatNumber(knownWordsSummary.totalUniqueWords)}
                </span>
                {knownWordPercent != null ? (
                  <span className="text-sm text-ctp-overlay2 ml-1">({knownWordPercent}%)</span>
                ) : null}
              </div>
            </div>
          </Tooltip>
        )}
      </div>
    </div>
  );
}
