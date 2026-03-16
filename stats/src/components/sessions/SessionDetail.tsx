import {
  ComposedChart,
  Area,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  ReferenceArea,
  ReferenceLine,
} from 'recharts';
import { useSessionDetail } from '../../hooks/useSessions';
import { CHART_THEME } from '../../lib/chart-theme';
import { EventType } from '../../types/stats';
import type { SessionEvent } from '../../types/stats';

interface SessionDetailProps {
  sessionId: number;
  cardsMined: number;
}

const tooltipStyle = {
  background: CHART_THEME.tooltipBg,
  border: `1px solid ${CHART_THEME.tooltipBorder}`,
  borderRadius: 6,
  color: CHART_THEME.tooltipText,
  fontSize: 11,
};

function formatTime(ms: number): string {
  return new Date(ms).toLocaleTimeString(undefined, {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

interface PauseRegion {
  startMs: number;
  endMs: number;
}

function buildPauseRegions(events: SessionEvent[]): PauseRegion[] {
  const regions: PauseRegion[] = [];
  const starts = events.filter((e) => e.eventType === EventType.PAUSE_START);
  const ends = events.filter((e) => e.eventType === EventType.PAUSE_END);

  for (const start of starts) {
    const end = ends.find((e) => e.tsMs > start.tsMs);
    regions.push({
      startMs: start.tsMs,
      endMs: end ? end.tsMs : start.tsMs + 2000,
    });
  }
  return regions;
}

interface ChartPoint {
  tsMs: number;
  activity: number;
  totalWords: number;
  paused: boolean;
}

export function SessionDetail({ sessionId, cardsMined }: SessionDetailProps) {
  const { timeline, events, loading, error } = useSessionDetail(sessionId);

  if (loading) return <div className="text-ctp-overlay2 text-xs p-2">Loading timeline...</div>;
  if (error) return <div className="text-ctp-red text-xs p-2">Error: {error}</div>;

  const sorted = [...timeline].reverse();
  const pauseRegions = buildPauseRegions(events);

  const chartData: ChartPoint[] = sorted.map((t, i) => {
    const prevWords = i > 0 ? sorted[i - 1]!.wordsSeen : 0;
    const delta = Math.max(0, t.wordsSeen - prevWords);
    const paused = pauseRegions.some((r) => t.sampleMs >= r.startMs && t.sampleMs <= r.endMs);
    return {
      tsMs: t.sampleMs,
      activity: delta,
      totalWords: t.wordsSeen,
      paused,
    };
  });

  const cardEvents = events.filter((e) => e.eventType === EventType.CARD_MINED);
  const seekEvents = events.filter(
    (e) => e.eventType === EventType.SEEK_FORWARD || e.eventType === EventType.SEEK_BACKWARD,
  );

  const pauseCount = events.filter((e) => e.eventType === EventType.PAUSE_START).length;
  const seekCount = seekEvents.length;
  const cardEventCount = cardEvents.length;

  const maxActivity = Math.max(...chartData.map((d) => d.activity), 1);
  const yMax = Math.ceil(maxActivity * 1.3);

  const tsMin = chartData.length > 0 ? chartData[0]!.tsMs : 0;
  const tsMax = chartData.length > 0 ? chartData[chartData.length - 1]!.tsMs : 0;

  return (
    <div className="bg-ctp-mantle border border-ctp-surface1 rounded-lg p-3 mt-1 space-y-3">
      {chartData.length > 0 && (
        <ResponsiveContainer width="100%" height={150}>
          <ComposedChart data={chartData} barCategoryGap={0} barGap={0}>
            <defs>
              <linearGradient id={`actGrad-${sessionId}`} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#c6a0f6" stopOpacity={0.5} />
                <stop offset="100%" stopColor="#c6a0f6" stopOpacity={0.05} />
              </linearGradient>
            </defs>
            <XAxis
              dataKey="tsMs"
              type="number"
              domain={[tsMin, tsMax]}
              tick={{ fontSize: 9, fill: CHART_THEME.tick }}
              axisLine={false}
              tickLine={false}
              tickFormatter={formatTime}
              interval="preserveStartEnd"
            />
            <YAxis
              yAxisId="left"
              tick={{ fontSize: 9, fill: CHART_THEME.tick }}
              axisLine={false}
              tickLine={false}
              width={24}
              domain={[0, yMax]}
              allowDecimals={false}
            />
            <YAxis
              yAxisId="right"
              orientation="right"
              tick={{ fontSize: 9, fill: CHART_THEME.tick }}
              axisLine={false}
              tickLine={false}
              width={30}
              allowDecimals={false}
            />
            <Tooltip
              contentStyle={tooltipStyle}
              labelFormatter={formatTime}
              formatter={(value: number, name: string) => {
                if (name === 'New words') return [`${value}`, 'New words'];
                if (name === 'Total words') return [`${value}`, 'Total words'];
                return [value, name];
              }}
            />

            {/* Pause shaded regions */}
            {pauseRegions.map((r, i) => (
              <ReferenceArea
                key={`pause-${i}`}
                yAxisId="left"
                x1={r.startMs}
                x2={r.endMs}
                y1={0}
                y2={yMax}
                fill="#f5a97f"
                fillOpacity={0.15}
                stroke="#f5a97f"
                strokeOpacity={0.4}
                strokeDasharray="3 3"
                strokeWidth={1}
              />
            ))}

            {/* Seek markers */}
            {seekEvents.map((e, i) => (
              <ReferenceLine
                key={`seek-${i}`}
                yAxisId="left"
                x={e.tsMs}
                stroke="#91d7e3"
                strokeWidth={1}
                strokeDasharray="3 4"
                strokeOpacity={0.5}
              />
            ))}

            {/* Card mined markers */}
            {cardEvents.map((e, i) => (
              <ReferenceLine
                key={`card-${i}`}
                yAxisId="left"
                x={e.tsMs}
                stroke="#a6da95"
                strokeWidth={2}
                strokeOpacity={0.8}
                label={{
                  value: '⛏',
                  position: 'top',
                  fill: '#a6da95',
                  fontSize: 14,
                  fontWeight: 700,
                }}
              />
            ))}

            <Area
              yAxisId="left"
              dataKey="activity"
              stroke="#c6a0f6"
              strokeWidth={1.5}
              fill={`url(#actGrad-${sessionId})`}
              name="New words"
              dot={false}
              activeDot={{ r: 3, fill: '#c6a0f6', stroke: '#1e2030', strokeWidth: 1 }}
              type="monotone"
              isAnimationActive={false}
            />
            <Line
              yAxisId="right"
              dataKey="totalWords"
              stroke="#8aadf4"
              strokeWidth={1.5}
              dot={false}
              activeDot={{ r: 3, fill: '#8aadf4', stroke: '#1e2030', strokeWidth: 1 }}
              name="Total words"
              type="monotone"
              isAnimationActive={false}
            />
          </ComposedChart>
        </ResponsiveContainer>
      )}

      <div className="flex flex-wrap items-center gap-4 text-[11px]">
        <span className="flex items-center gap-1.5">
          <span
            className="inline-block w-3 h-2 rounded-sm"
            style={{
              background:
                'linear-gradient(to bottom, rgba(198,160,246,0.5), rgba(198,160,246,0.05))',
            }}
          />
          <span className="text-ctp-overlay2">New words</span>
        </span>
        <span className="flex items-center gap-1.5">
          <span className="inline-block w-3 h-0.5 rounded" style={{ background: '#8aadf4' }} />
          <span className="text-ctp-overlay2">Total words</span>
        </span>
        {pauseCount > 0 && (
          <span className="flex items-center gap-1.5">
            <span
              className="inline-block w-3 h-2 rounded-sm"
              style={{
                background: 'rgba(245,169,127,0.2)',
                border: '1px solid rgba(245,169,127,0.5)',
              }}
            />
            <span className="text-ctp-overlay2">
              {pauseCount} pause{pauseCount !== 1 ? 's' : ''}
            </span>
          </span>
        )}
        {seekCount > 0 && (
          <span className="flex items-center gap-1.5">
            <span
              className="inline-block w-3 h-0.5 rounded"
              style={{ background: '#91d7e3', opacity: 0.7 }}
            />
            <span className="text-ctp-overlay2">
              {seekCount} seek{seekCount !== 1 ? 's' : ''}
            </span>
          </span>
        )}
        <span className="flex items-center gap-1.5">
          <span className="text-[12px]">⛏</span>
          <span className="text-ctp-green">
            {Math.max(cardEventCount, cardsMined)} card
            {Math.max(cardEventCount, cardsMined) !== 1 ? 's' : ''} mined
          </span>
        </span>
      </div>
    </div>
  );
}
