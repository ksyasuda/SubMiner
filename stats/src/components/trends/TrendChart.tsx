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

interface TrendChartProps {
  title: string;
  data: Array<{ label: string; value: number }>;
  color: string;
  type: 'bar' | 'line';
  formatter?: (value: number) => string;
  onBarClick?: (label: string) => void;
}

export function TrendChart({ title, data, color, type, formatter, onBarClick }: TrendChartProps) {
  const formatValue = (v: number) => (formatter ? [formatter(v), title] : [String(v), title]);

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-xs font-semibold text-ctp-text mb-2">{title}</h3>
      <ResponsiveContainer width="100%" height={CHART_DEFAULTS.height}>
        {type === 'bar' ? (
          <BarChart data={data} margin={CHART_DEFAULTS.margin}>
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
              tickFormatter={formatter}
            />
            <Tooltip contentStyle={TOOLTIP_CONTENT_STYLE} formatter={formatValue} />
            <Bar
              dataKey="value"
              fill={color}
              radius={[2, 2, 0, 0]}
              cursor={onBarClick ? 'pointer' : undefined}
              onClick={
                onBarClick ? (entry: { label: string }) => onBarClick(entry.label) : undefined
              }
            />
          </BarChart>
        ) : (
          <LineChart data={data} margin={CHART_DEFAULTS.margin}>
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
              tickFormatter={formatter}
            />
            <Tooltip contentStyle={TOOLTIP_CONTENT_STYLE} formatter={formatValue} />
            <Line dataKey="value" stroke={color} strokeWidth={2} dot={false} />
          </LineChart>
        )}
      </ResponsiveContainer>
    </div>
  );
}
