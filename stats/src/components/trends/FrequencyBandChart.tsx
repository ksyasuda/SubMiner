import { Bar, BarChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis } from 'recharts';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';
import { formatNumber } from '../../lib/formatters';
import {
  TREND_FREQUENCY_BANDS,
  type TrendFrequencyBand,
  type TrendFrequencyBandPoint,
} from '../../types/stats';

interface FrequencyBandChartProps {
  data: TrendFrequencyBandPoint[];
}

const BAND_LABELS: Record<TrendFrequencyBand, string> = {
  top1k: 'Top 1k',
  top5k: '1k to 5k',
  top10k: '5k to 10k',
  top20k: '10k to 20k',
  beyond20k: 'Beyond 20k',
  unranked: 'Unranked',
};

// Bands are ordered, so one hue runs from faint (common) to strong (rare): the bar
// brightens as new vocabulary moves further into rare words. Unranked words have
// no place in that order and get a neutral color.
const BAND_FILLS: Record<TrendFrequencyBand, { fill: string; fillOpacity: number }> = {
  top1k: { fill: '#8aadf4', fillOpacity: 0.2 },
  top5k: { fill: '#8aadf4', fillOpacity: 0.36 },
  top10k: { fill: '#8aadf4', fillOpacity: 0.54 },
  top20k: { fill: '#8aadf4', fillOpacity: 0.74 },
  beyond20k: { fill: '#8aadf4', fillOpacity: 1 },
  unranked: { fill: '#6e738d', fillOpacity: 0.6 },
};

interface BandTooltipEntry {
  dataKey?: unknown;
  value?: unknown;
}

function BandTooltip({
  active,
  label,
  payload,
}: {
  active?: boolean;
  label?: string | number;
  payload?: BandTooltipEntry[];
}) {
  if (!active || !payload?.length) return null;
  const rows = payload.map((entry) => ({
    band: TREND_FREQUENCY_BANDS.find((band) => band === entry.dataKey),
    count: typeof entry.value === 'number' ? entry.value : 0,
  }));
  const total = rows.reduce((sum, row) => sum + row.count, 0);
  return (
    <div style={{ ...TOOLTIP_CONTENT_STYLE, padding: '6px 10px' }}>
      <div style={{ color: CHART_THEME.tooltipLabel, marginBottom: 4, fontWeight: 600 }}>
        {label}: {formatNumber(total)} new words
      </div>
      {rows.map(({ band, count }) =>
        band ? (
          <div key={band} className="flex justify-between gap-4">
            <span>{BAND_LABELS[band]}</span>
            <span className="tabular-nums">
              {formatNumber(count)} ({total > 0 ? Math.round((count / total) * 100) : 0}%)
            </span>
          </div>
        ) : null,
      )}
    </div>
  );
}

export function FrequencyBandChart({ data }: FrequencyBandChartProps) {
  const rows = data.map((point) => ({ label: point.label, ...point.counts }));
  return (
    <div className="h-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
      <h3 className="mb-2 text-xs font-semibold text-ctp-text">New Words by Frequency Rank</h3>
      {rows.length === 0 ? (
        <div className="text-xs text-ctp-overlay2">No data</div>
      ) : (
        <>
          <ResponsiveContainer width="100%" height={CHART_DEFAULTS.height}>
            <BarChart data={rows} margin={CHART_DEFAULTS.margin} stackOffset="expand">
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
                width={40}
                tickFormatter={(value: number) => `${Math.round(value * 100)}%`}
              />
              <Tooltip content={<BandTooltip />} wrapperStyle={{ zIndex: 50 }} />
              {TREND_FREQUENCY_BANDS.map((band) => (
                <Bar key={band} dataKey={band} stackId="bands" {...BAND_FILLS[band]} />
              ))}
            </BarChart>
          </ResponsiveContainer>
          <div className="mt-2 flex flex-wrap gap-x-3 gap-y-1">
            {TREND_FREQUENCY_BANDS.map((band) => (
              <span key={band} className="flex items-center gap-1 text-[10px] text-ctp-subtext0">
                <span
                  className="inline-block h-2 w-2 shrink-0 rounded-sm"
                  style={{
                    backgroundColor: BAND_FILLS[band].fill,
                    opacity: BAND_FILLS[band].fillOpacity,
                  }}
                />
                {BAND_LABELS[band]}
              </span>
            ))}
          </div>
        </>
      )}
    </div>
  );
}
