import {
  BarChart,
  Bar,
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  ResponsiveContainer,
} from 'recharts';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';
import { formatCompactNumber, formatNumber } from '../../lib/formatters';

interface TrendChartProps {
  title: string;
  /** A null value is a gap: no data to plot, as opposed to a real zero. */
  data: Array<{ label: string; value: number | null }>;
  color: string;
  type: 'bar' | 'line';
  /** Formats axis ticks. Defaults to compact numbers such as 24K. */
  tickFormatter?: (value: number) => string;
  /** Formats the tooltip value. Defaults to the full number. */
  valueFormatter?: (value: number) => string;
  onBarClick?: (label: string, index: number) => void;
  height?: number;
}

const Y_AXIS_WIDTH = 40;

/** A point with gaps on both sides draws no line segment, so it gets a dot instead. */
export function isIsolatedPoint(values: ReadonlyArray<number | null>, index: number): boolean {
  return (
    values[index] !== null &&
    values[index] !== undefined &&
    (values[index - 1] ?? null) === null &&
    (values[index + 1] ?? null) === null
  );
}

/**
 * Leads a running total with zero. Each point is the total at the end of its period, so
 * without this the line would begin at the first day's whole total.
 */
export function startFromZero<T extends { label: string; value: number }>(
  points: T[],
): Array<T | { label: string; value: number }> {
  return points.length > 0 ? [{ label: 'Start', value: 0 }, ...points] : points;
}

export function TrendChart({
  title,
  data,
  color,
  type,
  tickFormatter = formatCompactNumber,
  valueFormatter = formatNumber,
  onBarClick,
  height = CHART_DEFAULTS.height,
}: TrendChartProps) {
  const formatValue = (value: unknown) =>
    typeof value === 'number' ? [valueFormatter(value), title] : ['Not enough data', title];
  const values = data.map((point) => point.value);

  const xAxis = (
    <XAxis
      dataKey="label"
      tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
      axisLine={{ stroke: CHART_THEME.axisLine }}
      tickLine={false}
    />
  );
  const yAxis = (
    <YAxis
      tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
      axisLine={{ stroke: CHART_THEME.axisLine }}
      tickLine={false}
      width={Y_AXIS_WIDTH}
      tickFormatter={tickFormatter}
    />
  );

  return (
    <div className="h-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-xs font-semibold text-ctp-text mb-2">{title}</h3>
      <ResponsiveContainer width="100%" height={height}>
        {type === 'bar' ? (
          <BarChart data={data} margin={CHART_DEFAULTS.margin}>
            <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
            {xAxis}
            {yAxis}
            <Tooltip contentStyle={TOOLTIP_CONTENT_STYLE} formatter={formatValue} />
            <Bar
              dataKey="value"
              fill={color}
              radius={[2, 2, 0, 0]}
              cursor={onBarClick ? 'pointer' : undefined}
              onClick={
                onBarClick
                  ? (entry: { label: string }, index: number) => onBarClick(entry.label, index)
                  : undefined
              }
            />
          </BarChart>
        ) : (
          <LineChart data={data} margin={CHART_DEFAULTS.margin}>
            <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
            {xAxis}
            {yAxis}
            <Tooltip contentStyle={TOOLTIP_CONTENT_STYLE} formatter={formatValue} />
            <Line
              dataKey="value"
              stroke={color}
              strokeWidth={2}
              dot={(props: { cx?: number; cy?: number; index?: number }) => (
                <circle
                  key={props.index}
                  cx={props.cx}
                  cy={props.cy}
                  r={props.index !== undefined && isIsolatedPoint(values, props.index) ? 2.5 : 0}
                  fill={color}
                />
              )}
            />
          </LineChart>
        )}
      </ResponsiveContainer>
    </div>
  );
}
