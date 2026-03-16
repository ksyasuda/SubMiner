import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { epochDayToDate } from '../../lib/formatters';

export interface PerAnimeDataPoint {
  epochDay: number;
  animeTitle: string;
  value: number;
}

interface StackedTrendChartProps {
  title: string;
  data: PerAnimeDataPoint[];
}

const LINE_COLORS = [
  '#8aadf4',
  '#c6a0f6',
  '#a6da95',
  '#f5a97f',
  '#f5bde6',
  '#91d7e3',
  '#ee99a0',
  '#f4dbd6',
];

function buildLineData(raw: PerAnimeDataPoint[]) {
  const totalByAnime = new Map<string, number>();
  for (const entry of raw) {
    totalByAnime.set(entry.animeTitle, (totalByAnime.get(entry.animeTitle) ?? 0) + entry.value);
  }

  const sorted = [...totalByAnime.entries()].sort((a, b) => b[1] - a[1]);
  const topTitles = sorted.slice(0, 7).map(([title]) => title);
  const topSet = new Set(topTitles);

  const byDay = new Map<number, Record<string, number>>();
  for (const entry of raw) {
    if (!topSet.has(entry.animeTitle)) continue;
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
      for (const title of topTitles) {
        row[title] = values[title] ?? 0;
      }
      return row;
    });

  return { points, seriesKeys: topTitles };
}

export function StackedTrendChart({ title, data }: StackedTrendChartProps) {
  const { points, seriesKeys } = buildLineData(data);

  const tooltipStyle = {
    background: '#363a4f',
    border: '1px solid #494d64',
    borderRadius: 6,
    color: '#cad3f5',
    fontSize: 12,
  };

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
      <ResponsiveContainer width="100%" height={120}>
        <AreaChart data={points}>
          <XAxis
            dataKey="label"
            tick={{ fontSize: 9, fill: '#a5adcb' }}
            axisLine={false}
            tickLine={false}
          />
          <YAxis
            tick={{ fontSize: 9, fill: '#a5adcb' }}
            axisLine={false}
            tickLine={false}
            width={28}
          />
          <Tooltip contentStyle={tooltipStyle} />
          {seriesKeys.map((key, i) => (
            <Area
              key={key}
              type="monotone"
              dataKey={key}
              stroke={LINE_COLORS[i % LINE_COLORS.length]}
              fill={LINE_COLORS[i % LINE_COLORS.length]}
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
              style={{ backgroundColor: LINE_COLORS[i % LINE_COLORS.length] }}
            />
            <span className="truncate">{key}</span>
          </span>
        ))}
      </div>
    </div>
  );
}
