import type { TrendSummaryTotals, TrendsDashboardData } from '../../types/stats';
import { formatCompactNumber, formatHours, formatNumber } from '../../lib/formatters';

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

function Tile({ label, value, detail }: { label: string; value: string; detail: string | null }) {
  return (
    <div className="min-w-0">
      <div className="text-[11px] text-ctp-subtext0">{label}</div>
      <div className="text-2xl font-semibold leading-tight text-ctp-text">{value}</div>
      {detail ? <div className="text-[11px] text-ctp-overlay2">{detail}</div> : null}
    </div>
  );
}

export function SummaryStrip({ summary, days }: SummaryStripProps) {
  const { current, previous } = summary;
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
          <Tile
            key={key}
            label={label}
            value={format(value)}
            detail={
              previousValue === null ? null : describePeriodChange(value, previousValue, days)
            }
          />
        );
      })}
      <Tile
        label="Active days"
        value={formatNumber(current.activeDays)}
        detail={previous ? `${formatNumber(previous.activeDays)} in previous ${days}` : null}
      />
    </section>
  );
}
