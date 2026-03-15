import {
  BarChart, Bar, LineChart, Line,
  XAxis, YAxis, Tooltip, ResponsiveContainer,
} from 'recharts';

interface TrendChartProps {
  title: string;
  data: Array<{ label: string; value: number }>;
  color: string;
  type: 'bar' | 'line';
  formatter?: (value: number) => string;
}

export function TrendChart({ title, data, color, type, formatter }: TrendChartProps) {
  const tooltipStyle = {
    background: '#363a4f', border: '1px solid #494d64', borderRadius: 6, color: '#cad3f5', fontSize: 12,
  };

  const formatValue = (v: number) => formatter ? [formatter(v), title] : [String(v), title];

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-xs font-semibold text-ctp-text mb-2">{title}</h3>
      <ResponsiveContainer width="100%" height={120}>
        {type === 'bar' ? (
          <BarChart data={data}>
            <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#a5adcb' }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fontSize: 9, fill: '#a5adcb' }} axisLine={false} tickLine={false} width={28} />
            <Tooltip contentStyle={tooltipStyle} formatter={formatValue} />
            <Bar dataKey="value" fill={color} radius={[2, 2, 0, 0]} />
          </BarChart>
        ) : (
          <LineChart data={data}>
            <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#a5adcb' }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fontSize: 9, fill: '#a5adcb' }} axisLine={false} tickLine={false} width={28} />
            <Tooltip contentStyle={tooltipStyle} formatter={formatValue} />
            <Line dataKey="value" stroke={color} strokeWidth={2} dot={false} />
          </LineChart>
        )}
      </ResponsiveContainer>
    </div>
  );
}
