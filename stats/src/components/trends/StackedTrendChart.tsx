import {
  AreaChart,
  Area,
  CartesianGrid,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
} from 'recharts';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';
import { epochDayToDate } from '../../lib/formatters';

export interface PerAnimeDataPoint {
  epochDay: number;
  animeTitle: string;
  value: number;
}

interface StackedTrendChartProps {
  title: string;
  data: PerAnimeDataPoint[];
  colorPalette?: string[];
}

const DEFAULT_LINE_COLORS = [
  '#8aadf4',
  '#c6a0f6',
  '#a6da95',
  '#f5a97f',
  '#f5bde6',
  '#91d7e3',
  '#ee99a0',
  '#f4dbd6',
];

export function buildLineData(raw: PerAnimeDataPoint[]) {
  const totalByAnime = new Map<string, number>();
  for (const entry of raw) {
    totalByAnime.set(entry.animeTitle, (totalByAnime.get(entry.animeTitle) ?? 0) + entry.value);
  }

  const seriesKeys = [...totalByAnime.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .map(([title]) => title);

  const byDay = new Map<number, Record<string, number>>();
  for (const entry of raw) {
    const row = byDay.get(entry.epochDay) ?? {};
    row[entry.animeTitle] = (row[entry.animeTitle] ?? 0) + Math.round(entry.value * 10) / 10;
    byDay.set(entry.epochDay, row);
  }

  const points = [...byDay.entries()]
    .sort(([a], [b]) => a - b)
    .map(([epochDay, values]) => {
      const row: Record<string, string | number> = {
        label: epochDayToDate(epochDay).toLocaleDateString(undefined, {
          month: 'short',
          day: 'numeric',
        }),
      };
      for (const title of seriesKeys) {
        row[title] = values[title] ?? 0;
      }
      return row;
    });

  return { points, seriesKeys };
}

export function StackedTrendChart({ title, data, colorPalette }: StackedTrendChartProps) {
  const { points, seriesKeys } = buildLineData(data);
  const colors = colorPalette ?? DEFAULT_LINE_COLORS;

  if (points.length === 0) {
    return (
      <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
        <h3 className="text-xs font-semibold text-ctp-text mb-2">{title}</h3>
        <div className="text-xs text-ctp-overlay2">No data</div>
      </div>
    );
  }

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-xs font-semibold text-ctp-text mb-2">{title}</h3>
      <ResponsiveContainer width="100%" height={CHART_DEFAULTS.height}>
        <AreaChart data={points} margin={CHART_DEFAULTS.margin}>
          <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
          <XAxis
            dataKey="label"
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
          <Tooltip contentStyle={TOOLTIP_CONTENT_STYLE} />
          {seriesKeys.map((key, i) => (
            <Area
              key={key}
              type="monotone"
              dataKey={key}
              stroke={colors[i % colors.length]}
              fill={colors[i % colors.length]}
              fillOpacity={0.15}
              strokeWidth={1.5}
              connectNulls
            />
          ))}
        </AreaChart>
      </ResponsiveContainer>
      <div className="flex flex-wrap gap-x-3 gap-y-1 mt-2 overflow-hidden max-h-10">
        {seriesKeys.map((key, i) => (
          <span
            key={key}
            className="flex items-center gap-1 text-[10px] text-ctp-subtext0 max-w-[140px]"
            title={key}
          >
            <span
              className="inline-block w-2 h-2 rounded-full shrink-0"
              style={{ backgroundColor: colors[i % colors.length] }}
            />
            <span className="truncate">{key}</span>
          </span>
        ))}
      </div>
    </div>
  );
}
