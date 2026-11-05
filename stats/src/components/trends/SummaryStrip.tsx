import type { TrendSummaryTotals, TrendsDashboardData } from '../../types/stats';
import { formatCompactNumber, formatHours, formatNumber } from '../../lib/formatters';
import { SummaryTile } from '../layout/SummaryTile';

interface SummaryStripProps {
  summary: TrendsDashboardData['summary'];
  days: number;
}

type NumericKey = Exclude<keyof TrendSummaryTotals, 'activeDays'>;

const TILES: Array<{ key: NumericKey; label: string; format: (value: number) => string }> = [
  { key: 'watchTimeMin', label: 'Watch time', format: formatHours },
  { key: 'words', label: 'Words seen', format: formatCompactNumber },
  { key: 'newWords', label: 'New words', format: formatNumber },
  { key: 'cards', label: 'Cards mined', format: formatNumber },
  { key: 'episodes', label: 'Episodes finished', format: formatNumber },
];

/** Compares a total with the same-length period before it; null when both are zero. */
export function describePeriodChange(
  current: number,
  previous: number,
  days: number,
): string | null {
  const period = `previous ${days} day${days === 1 ? '' : 's'}`;
  if (previous === 0) return current === 0 ? null : `None in ${period}`;
  if (current === previous) return `Same as ${period}`;
  const percent = Math.round(((current - previous) / previous) * 100);
  if (percent === 0) return `Under 1% from ${period}`;
  return `${percent > 0 ? '+' : '\u2212'}${Math.abs(percent)}% vs ${period}`;
}

function isEmptyPeriod(totals: TrendSummaryTotals): boolean {
  return totals.activeDays === 0 && TILES.every(({ key }) => !totals[key]);
}

export function SummaryStrip({ summary, days }: SummaryStripProps) {
  const { current } = summary;
  // An empty period gets one note, because the same sentence on every tile reads as missing data.
  const previousIsEmpty = summary.previous !== null && isEmptyPeriod(summary.previous);
  const previous = previousIsEmpty ? null : summary.previous;
  return (
    <section
      aria-label="Period summary"
      className="grid grid-cols-2 gap-4 rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4 sm:grid-cols-3 lg:grid-cols-6"
    >
      {TILES.map(({ key, label, format }) => {
        const value = current[key];
        if (value === null) return null;
        const previousValue = previous?.[key] ?? null;
        return (
          <SummaryTile
            key={key}
            label={label}
            value={format(value)}
            detail={
              previousValue === null ? null : describePeriodChange(value, previousValue, days)
            }
          />
        );
      })}
      <SummaryTile
        label="Active days"
        value={formatNumber(current.activeDays)}
        detail={previous ? `${formatNumber(previous.activeDays)} in previous ${days}` : null}
      />
      {previousIsEmpty ? (
        <p className="col-span-full text-[11px] text-ctp-overlay2">
          Not watched in the previous {days} day{days === 1 ? '' : 's'}
        </p>
      ) : null}
    </section>
  );
}
