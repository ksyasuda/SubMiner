import { useState } from 'react';
import { useTranslation } from '../../i18n';
import { BarChart, Bar, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { epochDayToDate } from '../../lib/formatters';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';
import type { DailyRollup } from '../../types/stats';

interface WatchTimeChartProps {
  rollups: DailyRollup[];
}

type Range = 14 | 30 | 90;

export function WatchTimeChart({ rollups }: WatchTimeChartProps) {
  const { t } = useTranslation();
  const [range, setRange] = useState<Range>(14);

  const formatActiveMinutes = (value: number | string, _name?: string, _payload?: unknown) => {
    const minutes = Number(value);
    return [
      `${Number.isFinite(minutes) ? minutes : 0} ${t('stats.watchTime.minutes')}`,
      t('stats.watchTime.activeTime'),
    ];
  };

  const byDay = new Map<number, number>();
  for (const r of rollups) {
    byDay.set(r.rollupDayOrMonth, (byDay.get(r.rollupDayOrMonth) ?? 0) + r.totalActiveMin);
  }
  const chartData = Array.from(byDay.entries())
    .sort(([dayA], [dayB]) => dayA - dayB)
    .map(([day, mins]) => ({
      date: epochDayToDate(day).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }),
      minutes: Math.round(mins),
    }))
    .slice(-range);

  const ranges: Range[] = [14, 30, 90];

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-ctp-text">{t('stats.watchTime.title')}</h3>
        <div className="flex bg-ctp-surface0 rounded-lg p-0.5 border border-ctp-surface1">
          {ranges.map((r) => (
            <button
              key={r}
              onClick={() => setRange(r)}
              className={`px-2.5 py-1 text-xs rounded-md transition-colors ${
                range === r
                  ? 'bg-ctp-surface2 text-ctp-text shadow-sm'
                  : 'text-ctp-overlay2 hover:text-ctp-subtext0'
              }`}
            >
              {r}d
            </button>
          ))}
        </div>
      </div>
      <ResponsiveContainer width="100%" height={CHART_DEFAULTS.height}>
        <BarChart data={chartData} margin={CHART_DEFAULTS.margin}>
          <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
          <XAxis
            dataKey="date"
            tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
            axisLine={{ stroke: CHART_THEME.axisLine }}
            tickLine={false}
          />
          <YAxis
            tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
            axisLine={{ stroke: CHART_THEME.axisLine }}
            tickLine={false}
            width={32}
          />
          <Tooltip
            contentStyle={TOOLTIP_CONTENT_STYLE}
            labelStyle={{ color: CHART_THEME.tooltipLabel }}
            formatter={formatActiveMinutes}
          />
          <Bar dataKey="minutes" fill={CHART_THEME.barFill} radius={[3, 3, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
