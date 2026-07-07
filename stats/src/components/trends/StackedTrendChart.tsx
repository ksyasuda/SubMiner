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
  maxSeries?: number | null;
  maxSeriesMode?: SeriesRankMode;
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

// Wrap the tooltip into extra columns once a single column would get too tall,
// keeping it compact instead of overflowing behind the charts below it.
const TOOLTIP_ROWS_PER_COLUMN = 8;
const TOOLTIP_MAX_COLUMNS = 3;

export function tooltipColumnCount(itemCount: number): number {
  const columns = Math.ceil(itemCount / TOOLTIP_ROWS_PER_COLUMN);
  return Math.min(TOOLTIP_MAX_COLUMNS, Math.max(1, columns));
}

interface TooltipEntry {
  name?: string | number;
  value?: string | number;
  color?: string;
}

interface StackedTooltipProps {
  active?: boolean;
  label?: string | number;
  payload?: TooltipEntry[];
}

function StackedTooltip({ active, label, payload }: StackedTooltipProps) {
  if (!active || !payload || payload.length === 0) {
    return null;
  }
  const columns = tooltipColumnCount(payload.length);
  return (
    <div
      style={{
        ...TOOLTIP_CONTENT_STYLE,
        padding: '6px 10px',
        maxWidth: '80vw',
      }}
    >
      {label !== undefined && (
        <div style={{ color: CHART_THEME.tooltipLabel, marginBottom: 4, fontWeight: 600 }}>
          {label}
        </div>
      )}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))`,
          columnGap: 12,
          rowGap: 2,
        }}
      >
        {payload.map((entry, index) => (
          <div
            key={`${entry.name ?? index}`}
            style={{ display: 'flex', alignItems: 'center', gap: 6, whiteSpace: 'nowrap' }}
          >
            <span
              style={{
                display: 'inline-block',
                width: 8,
                height: 8,
                borderRadius: 9999,
                background: entry.color,
                flexShrink: 0,
              }}
            />
            <span
              style={{
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                maxWidth: 220,
              }}
            >
              {entry.name}
            </span>
            <span style={{ marginLeft: 'auto', color: CHART_THEME.tooltipLabel }}>
              {entry.value}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

export type SeriesRankMode = 'total' | 'recent';

// Per-title ranking key. `total` sums the values; `recentDay` is the latest
// epoch day the title's (cumulative) value increased, i.e. its last day of
// real activity — used to keep the most recently watched titles.
function rankTitles(raw: PerAnimeDataPoint[]): Map<string, { total: number; recentDay: number }> {
  const pointsByTitle = new Map<string, PerAnimeDataPoint[]>();
  for (const entry of raw) {
    const list = pointsByTitle.get(entry.animeTitle) ?? [];
    list.push(entry);
    pointsByTitle.set(entry.animeTitle, list);
  }

  const stats = new Map<string, { total: number; recentDay: number }>();
  for (const [title, points] of pointsByTitle) {
    const sorted = [...points].sort((a, b) => a.epochDay - b.epochDay);
    let total = 0;
    let previous = 0;
    let recentDay = Number.NEGATIVE_INFINITY;
    for (const point of sorted) {
      total += point.value;
      if (point.value > previous) {
        recentDay = point.epochDay;
      }
      previous = point.value;
    }
    stats.set(title, { total, recentDay });
  }
  return stats;
}

export function buildLineData(
  raw: PerAnimeDataPoint[],
  maxSeries?: number | null,
  mode: SeriesRankMode = 'total',
) {
  const stats = rankTitles(raw);

  let seriesKeys = [...stats.entries()]
    .sort((a, b) => {
      if (mode === 'recent') {
        return (
          b[1].recentDay - a[1].recentDay || b[1].total - a[1].total || a[0].localeCompare(b[0])
        );
      }
      return b[1].total - a[1].total || a[0].localeCompare(b[0]);
    })
    .map(([title]) => title);
  if (typeof maxSeries === 'number' && maxSeries > 0) {
    seriesKeys = seriesKeys.slice(0, maxSeries);
  }
  const seriesSet = new Set(seriesKeys);

  const byDay = new Map<number, Record<string, number>>();
  for (const entry of raw) {
    if (!seriesSet.has(entry.animeTitle)) continue;
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

export function StackedTrendChart({
  title,
  data,
  colorPalette,
  maxSeries,
  maxSeriesMode,
}: StackedTrendChartProps) {
  const { points, seriesKeys } = buildLineData(data, maxSeries, maxSeriesMode);
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
          <Tooltip
            content={<StackedTooltip />}
            wrapperStyle={{ zIndex: 50 }}
            allowEscapeViewBox={{ x: false, y: true }}
          />
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
