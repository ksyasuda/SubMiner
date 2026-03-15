import { useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { epochDayToDate } from '../../lib/formatters';
import { CHART_THEME } from '../../lib/chart-theme';
import type { DailyRollup } from '../../types/stats';

interface MediaWatchChartProps {
  rollups: DailyRollup[];
}

type Range = 14 | 30 | 90;

function formatActiveMinutes(value: number | string) {
  const minutes = Number(value);
  return [`${Number.isFinite(minutes) ? minutes : 0} min`, 'Active Time'];
}

export function MediaWatchChart({ rollups }: MediaWatchChartProps) {
  const [range, setRange] = useState<Range>(30);

  const byDay = new Map<number, number>();
  for (const r of rollups) {
    byDay.set(r.rollupDayOrMonth, (byDay.get(r.rollupDayOrMonth) ?? 0) + r.totalActiveMin);
  }
  const chartData = Array.from(byDay.entries())
    .sort(([a], [b]) => a - b)
    .map(([day, mins]) => ({
      date: epochDayToDate(day).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }),
      minutes: Math.round(mins),
    }))
    .slice(-range);

  const ranges: Range[] = [14, 30, 90];

  if (chartData.length === 0) {
    return null;
  }

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-ctp-text">Watch Time</h3>
        <div className="flex gap-1">
          {ranges.map((r) => (
            <button
              key={r}
              onClick={() => setRange(r)}
              className={`px-2 py-0.5 text-xs rounded ${
                range === r
                  ? 'bg-ctp-surface2 text-ctp-text'
                  : 'text-ctp-overlay2 hover:text-ctp-subtext0'
              }`}
            >
              {r}d
            </button>
          ))}
        </div>
      </div>
      <ResponsiveContainer width="100%" height={160}>
        <BarChart data={chartData}>
          <XAxis dataKey="date" tick={{ fontSize: 10, fill: CHART_THEME.tick }} axisLine={false} tickLine={false} />
          <YAxis tick={{ fontSize: 10, fill: CHART_THEME.tick }} axisLine={false} tickLine={false} width={30} />
          <Tooltip
            contentStyle={{
              background: CHART_THEME.tooltipBg,
              border: `1px solid ${CHART_THEME.tooltipBorder}`,
              borderRadius: 6,
              color: CHART_THEME.tooltipText,
              fontSize: 12,
            }}
            labelStyle={{ color: CHART_THEME.tooltipLabel }}
            formatter={formatActiveMinutes}
          />
          <Bar dataKey="minutes" fill={CHART_THEME.barFill} radius={[3, 3, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
