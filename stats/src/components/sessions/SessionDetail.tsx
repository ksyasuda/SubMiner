import {
  LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer,
  ReferenceLine,
} from 'recharts';
import { useSessionDetail } from '../../hooks/useSessions';
import { CHART_THEME } from '../../lib/chart-theme';
import { EventType } from '../../types/stats';

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

const EVENT_COLORS: Partial<Record<number, { color: string; label: string }>> = {
  [EventType.CARD_MINED]: { color: '#a6da95', label: 'Card mined' },
  [EventType.PAUSE_START]: { color: '#f5a97f', label: 'Pause' },
};

export function SessionDetail({ sessionId, cardsMined }: SessionDetailProps) {
  const { timeline, events, loading, error } = useSessionDetail(sessionId);

  if (loading) return <div className="text-ctp-overlay2 text-xs p-2">Loading timeline...</div>;
  if (error) return <div className="text-ctp-red text-xs p-2">Error: {error}</div>;

  const chartData = [...timeline]
    .reverse()
    .map((t) => ({
      tsMs: t.sampleMs,
      time: formatTime(t.sampleMs),
      words: t.wordsSeen,
      cards: t.cardsMined,
    }));

  const pauseCount = events.filter((e) => e.eventType === EventType.PAUSE_START).length;
  const seekCount = events.filter(
    (e) => e.eventType === EventType.SEEK_FORWARD || e.eventType === EventType.SEEK_BACKWARD,
  ).length;
  const cardEventCount = events.filter((e) => e.eventType === EventType.CARD_MINED).length;

  const markerEvents = events.filter((e) => EVENT_COLORS[e.eventType]);

  return (
    <div className="bg-ctp-mantle border border-ctp-surface1 rounded-lg p-3 mt-1 space-y-3">
      {chartData.length > 0 && (
        <ResponsiveContainer width="100%" height={120}>
          <LineChart data={chartData}>
            <XAxis
              dataKey="time"
              tick={{ fontSize: 9, fill: CHART_THEME.tick }}
              axisLine={false}
              tickLine={false}
            />
            <YAxis
              tick={{ fontSize: 9, fill: CHART_THEME.tick }}
              axisLine={false}
              tickLine={false}
              width={28}
            />
            <Tooltip contentStyle={tooltipStyle} />
            <Line dataKey="words" stroke="#c6a0f6" strokeWidth={1.5} dot={false} name="Words" />
            <Line dataKey="cards" stroke="#a6da95" strokeWidth={1.5} dot={false} name="Cards" />
            {markerEvents.map((e, i) => {
              const cfg = EVENT_COLORS[e.eventType]!;
              const matchIdx = chartData.findIndex((d) => d.tsMs >= e.tsMs);
              const x = matchIdx >= 0 ? chartData[matchIdx]!.time : null;
              if (!x) return null;
              return (
                <ReferenceLine
                  key={`${e.eventType}-${i}`}
                  x={x}
                  stroke={cfg.color}
                  strokeDasharray="3 3"
                  strokeOpacity={0.6}
                  label=""
                />
              );
            })}
          </LineChart>
        </ResponsiveContainer>
      )}

      <div className="flex flex-wrap gap-4 text-xs text-ctp-subtext0">
        <span>{pauseCount} pause{pauseCount !== 1 ? 's' : ''}</span>
        <span>{seekCount} seek{seekCount !== 1 ? 's' : ''}</span>
        <span className="text-ctp-green">{Math.max(cardEventCount, cardsMined)} card{Math.max(cardEventCount, cardsMined) !== 1 ? 's' : ''} mined</span>
      </div>

      {markerEvents.length > 0 && (
        <div className="flex flex-wrap gap-3 text-[10px]">
          {Object.entries(EVENT_COLORS).map(([type, cfg]) => {
            if (!cfg) return null;
            const count = markerEvents.filter((e) => e.eventType === Number(type)).length;
            if (count === 0) return null;
            return (
              <span key={type} className="flex items-center gap-1">
                <span className="inline-block w-2.5 h-0.5 rounded" style={{ background: cfg.color }} />
                <span className="text-ctp-overlay2">{cfg.label} ({count})</span>
              </span>
            );
          })}
        </div>
      )}
    </div>
  );
}
