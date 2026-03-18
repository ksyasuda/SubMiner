import {
  AreaChart,
  Area,
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  ReferenceArea,
  ReferenceLine,
  CartesianGrid,
} from 'recharts';
import { useSessionDetail } from '../../hooks/useSessions';
import type { KnownWordsTimelinePoint } from '../../hooks/useSessions';
import { CHART_THEME } from '../../lib/chart-theme';
import { buildLookupRateDisplay, getYomitanLookupEvents } from '../../lib/yomitan-lookup';
import { getSessionDisplayWordCount } from '../../lib/session-word-count';
import { EventType } from '../../types/stats';
import type { SessionEvent, SessionSummary } from '../../types/stats';

interface SessionDetailProps {
  session: SessionSummary;
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

/** Build a lookup: linesSeen → knownWordsSeen */
function buildKnownWordsLookup(
  knownWordsTimeline: KnownWordsTimelinePoint[],
): Map<number, number> {
  const map = new Map<number, number>();
  for (const pt of knownWordsTimeline) {
    map.set(pt.linesSeen, pt.knownWordsSeen);
  }
  return map;
}

/** For a given linesSeen value, find the closest known words count (floor lookup). */
function lookupKnownWords(map: Map<number, number>, linesSeen: number): number {
  if (map.size === 0) return 0;
  if (map.has(linesSeen)) return map.get(linesSeen)!;
  let best = 0;
  for (const k of map.keys()) {
    if (k <= linesSeen && k > best) {
      best = k;
    }
  }
  return best > 0 ? map.get(best)! : 0;
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

interface RatioChartPoint {
  tsMs: number;
  knownPct: number;
  unknownPct: number;
  knownWords: number;
  unknownWords: number;
  totalWords: number;
}

interface FallbackChartPoint {
  tsMs: number;
  totalWords: number;
}

type TimelineEntry = {
  sampleMs: number;
  linesSeen: number;
  wordsSeen: number;
  tokensSeen: number;
};

export function SessionDetail({ session }: SessionDetailProps) {
  const { timeline, events, knownWordsTimeline, loading, error } = useSessionDetail(
    session.sessionId,
  );

  if (loading) return <div className="text-ctp-overlay2 text-xs p-2">Loading timeline...</div>;
  if (error) return <div className="text-ctp-red text-xs p-2">Error: {error}</div>;

  const sorted = [...timeline].reverse();
  const knownWordsMap = buildKnownWordsLookup(knownWordsTimeline);
  const hasKnownWords = knownWordsMap.size > 0;

  const cardEvents = events.filter((e) => e.eventType === EventType.CARD_MINED);
  const seekEvents = events.filter(
    (e) => e.eventType === EventType.SEEK_FORWARD || e.eventType === EventType.SEEK_BACKWARD,
  );
  const yomitanLookupEvents = getYomitanLookupEvents(events);
  const lookupRate = buildLookupRateDisplay(
    session.yomitanLookupCount,
    getSessionDisplayWordCount(session),
  );
  const pauseCount = events.filter((e) => e.eventType === EventType.PAUSE_START).length;
  const seekCount = seekEvents.length;
  const cardEventCount = cardEvents.length;
  const pauseRegions = buildPauseRegions(events);

  if (hasKnownWords) {
    return (
      <RatioView
        sorted={sorted}
        knownWordsMap={knownWordsMap}
        cardEvents={cardEvents}
        yomitanLookupEvents={yomitanLookupEvents}
        pauseRegions={pauseRegions}
        pauseCount={pauseCount}
        seekCount={seekCount}
        cardEventCount={cardEventCount}
        lookupRate={lookupRate}
        session={session}
      />
    );
  }

  return (
    <FallbackView
      sorted={sorted}
      cardEvents={cardEvents}
      yomitanLookupEvents={yomitanLookupEvents}
      pauseRegions={pauseRegions}
      pauseCount={pauseCount}
      seekCount={seekCount}
      cardEventCount={cardEventCount}
      lookupRate={lookupRate}
      session={session}
    />
  );
}

/* ── Ratio View (primary design) ────────────────────────────────── */

function RatioView({
  sorted,
  knownWordsMap,
  cardEvents,
  yomitanLookupEvents,
  pauseRegions,
  pauseCount,
  seekCount,
  cardEventCount,
  lookupRate,
  session,
}: {
  sorted: TimelineEntry[];
  knownWordsMap: Map<number, number>;
  cardEvents: SessionEvent[];
  yomitanLookupEvents: SessionEvent[];
  pauseRegions: PauseRegion[];
  pauseCount: number;
  seekCount: number;
  cardEventCount: number;
  lookupRate: ReturnType<typeof buildLookupRateDisplay>;
  session: SessionSummary;
}) {
  const chartData: RatioChartPoint[] = [];
  for (const t of sorted) {
    const totalWords = getSessionDisplayWordCount(t);
    if (totalWords === 0) continue;
    const knownWords = Math.min(lookupKnownWords(knownWordsMap, t.linesSeen), totalWords);
    const unknownWords = totalWords - knownWords;
    const knownPct = (knownWords / totalWords) * 100;
    chartData.push({
      tsMs: t.sampleMs,
      knownPct,
      unknownPct: 100 - knownPct,
      knownWords,
      unknownWords,
      totalWords,
    });
  }

  if (chartData.length === 0) {
    return <div className="text-ctp-overlay2 text-xs p-2">No word data for this session.</div>;
  }

  const tsMin = chartData[0]!.tsMs;
  const tsMax = chartData[chartData.length - 1]!.tsMs;
  const finalTotal = chartData[chartData.length - 1]!.totalWords;

  const sparkData = chartData.map((d) => ({ tsMs: d.tsMs, totalWords: d.totalWords }));

  return (
    <div className="bg-ctp-mantle border border-ctp-surface1 rounded-lg p-3 mt-1 space-y-1">
      {/* ── Top: Percentage area chart ── */}
      <ResponsiveContainer width="100%" height={130}>
        <AreaChart data={chartData}>
          <defs>
            <linearGradient id={`knownGrad-${session.sessionId}`} x1="0" y1="1" x2="0" y2="0">
              <stop offset="0%" stopColor="#a6da95" stopOpacity={0.4} />
              <stop offset="100%" stopColor="#a6da95" stopOpacity={0.15} />
            </linearGradient>
            <linearGradient id={`unknownGrad-${session.sessionId}`} x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#c6a0f6" stopOpacity={0.25} />
              <stop offset="100%" stopColor="#c6a0f6" stopOpacity={0.08} />
            </linearGradient>
          </defs>

          <CartesianGrid
            horizontal
            vertical={false}
            stroke="#494d64"
            strokeDasharray="4 4"
            strokeOpacity={0.4}
          />

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
            yAxisId="pct"
            orientation="right"
            domain={[0, 100]}
            ticks={[0, 50, 100]}
            tick={{ fontSize: 9, fill: CHART_THEME.tick }}
            tickFormatter={(v: number) => `${v}%`}
            axisLine={false}
            tickLine={false}
            width={32}
          />

          <Tooltip
            contentStyle={tooltipStyle}
            labelFormatter={formatTime}
            formatter={(_value: number, name: string, props: { payload?: RatioChartPoint }) => {
              const d = props.payload;
              if (!d) return [_value, name];
              if (name === 'Known')
                return [`${d.knownWords.toLocaleString()} (${d.knownPct.toFixed(1)}%)`, 'Known'];
              if (name === 'Unknown')
                return [
                  `${d.unknownWords.toLocaleString()} (${d.unknownPct.toFixed(1)}%)`,
                  'Unknown',
                ];
              return [_value, name];
            }}
            itemSorter={() => -1}
          />

          {/* Pause shaded regions */}
          {pauseRegions.map((r, i) => (
            <ReferenceArea
              key={`pause-${i}`}
              yAxisId="pct"
              x1={r.startMs}
              x2={r.endMs}
              y1={0}
              y2={100}
              fill="#f5a97f"
              fillOpacity={0.15}
              stroke="#f5a97f"
              strokeOpacity={0.4}
              strokeDasharray="3 3"
              strokeWidth={1}
            />
          ))}

          {/* Card mine markers */}
          {cardEvents.map((e, i) => (
            <ReferenceLine
              key={`card-${i}`}
              yAxisId="pct"
              x={e.tsMs}
              stroke="#a6da95"
              strokeWidth={2}
              strokeOpacity={0.8}
              label={{
                value: '\u26CF',
                position: 'top',
                fill: '#a6da95',
                fontSize: 14,
                fontWeight: 700,
              }}
            />
          ))}

          {/* Yomitan lookup markers */}
          {yomitanLookupEvents.map((e, i) => (
            <ReferenceLine
              key={`yomitan-${i}`}
              yAxisId="pct"
              x={e.tsMs}
              stroke="#b7bdf8"
              strokeWidth={1.5}
              strokeDasharray="2 3"
              strokeOpacity={0.7}
            />
          ))}

          <Area
            yAxisId="pct"
            dataKey="knownPct"
            stackId="ratio"
            stroke="#a6da95"
            strokeWidth={1.5}
            fill={`url(#knownGrad-${session.sessionId})`}
            name="Known"
            type="monotone"
            dot={false}
            activeDot={{ r: 3, fill: '#a6da95', stroke: '#1e2030', strokeWidth: 1 }}
            isAnimationActive={false}
          />
          <Area
            yAxisId="pct"
            dataKey="unknownPct"
            stackId="ratio"
            stroke="#c6a0f6"
            strokeWidth={0}
            fill={`url(#unknownGrad-${session.sessionId})`}
            name="Unknown"
            type="monotone"
            isAnimationActive={false}
          />
        </AreaChart>
      </ResponsiveContainer>

      {/* ── Bottom: Word accumulation sparkline ── */}
      <div className="flex items-center gap-2 border-t border-ctp-surface1 pt-1">
        <span className="text-[9px] text-ctp-overlay0 whitespace-nowrap">total words</span>
        <div className="flex-1 h-[28px]">
          <ResponsiveContainer width="100%" height={28}>
            <LineChart data={sparkData}>
              <XAxis dataKey="tsMs" type="number" domain={[tsMin, tsMax]} hide />
              <YAxis hide />
              <Line
                dataKey="totalWords"
                stroke="#8aadf4"
                strokeWidth={1.5}
                strokeOpacity={0.8}
                dot={false}
                type="monotone"
                isAnimationActive={false}
              />
            </LineChart>
          </ResponsiveContainer>
        </div>
        <span className="text-[10px] text-ctp-blue font-semibold whitespace-nowrap tabular-nums">
          {finalTotal.toLocaleString()}
        </span>
      </div>

      {/* ── Stats bar ── */}
      <StatsBar
        hasKnownWords
        pauseCount={pauseCount}
        seekCount={seekCount}
        cardEventCount={cardEventCount}
        session={session}
        lookupRate={lookupRate}
      />
    </div>
  );
}

/* ── Fallback View (no known words data) ────────────────────────── */

function FallbackView({
  sorted,
  cardEvents,
  yomitanLookupEvents,
  pauseRegions,
  pauseCount,
  seekCount,
  cardEventCount,
  lookupRate,
  session,
}: {
  sorted: TimelineEntry[];
  cardEvents: SessionEvent[];
  yomitanLookupEvents: SessionEvent[];
  pauseRegions: PauseRegion[];
  pauseCount: number;
  seekCount: number;
  cardEventCount: number;
  lookupRate: ReturnType<typeof buildLookupRateDisplay>;
  session: SessionSummary;
}) {
  const chartData: FallbackChartPoint[] = [];
  for (const t of sorted) {
    const totalWords = getSessionDisplayWordCount(t);
    if (totalWords === 0) continue;
    chartData.push({ tsMs: t.sampleMs, totalWords });
  }

  if (chartData.length === 0) {
    return <div className="text-ctp-overlay2 text-xs p-2">No word data for this session.</div>;
  }

  const tsMin = chartData[0]!.tsMs;
  const tsMax = chartData[chartData.length - 1]!.tsMs;

  return (
    <div className="bg-ctp-mantle border border-ctp-surface1 rounded-lg p-3 mt-1 space-y-3">
      <ResponsiveContainer width="100%" height={130}>
        <LineChart data={chartData}>
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
            tick={{ fontSize: 9, fill: CHART_THEME.tick }}
            axisLine={false}
            tickLine={false}
            width={30}
            allowDecimals={false}
          />
          <Tooltip
            contentStyle={tooltipStyle}
            labelFormatter={formatTime}
            formatter={(value: number) => [`${value.toLocaleString()}`, 'Total words']}
          />

          {pauseRegions.map((r, i) => (
            <ReferenceArea
              key={`pause-${i}`}
              x1={r.startMs}
              x2={r.endMs}
              fill="#f5a97f"
              fillOpacity={0.15}
              stroke="#f5a97f"
              strokeOpacity={0.4}
              strokeDasharray="3 3"
              strokeWidth={1}
            />
          ))}

          {cardEvents.map((e, i) => (
            <ReferenceLine
              key={`card-${i}`}
              x={e.tsMs}
              stroke="#a6da95"
              strokeWidth={2}
              strokeOpacity={0.8}
              label={{
                value: '\u26CF',
                position: 'top',
                fill: '#a6da95',
                fontSize: 14,
                fontWeight: 700,
              }}
            />
          ))}
          {yomitanLookupEvents.map((e, i) => (
            <ReferenceLine
              key={`yomitan-${i}`}
              x={e.tsMs}
              stroke="#b7bdf8"
              strokeWidth={1.5}
              strokeDasharray="2 3"
              strokeOpacity={0.7}
            />
          ))}

          <Line
            dataKey="totalWords"
            stroke="#8aadf4"
            strokeWidth={1.5}
            dot={false}
            activeDot={{ r: 3, fill: '#8aadf4', stroke: '#1e2030', strokeWidth: 1 }}
            name="Total words"
            type="monotone"
            isAnimationActive={false}
          />
        </LineChart>
      </ResponsiveContainer>

      <StatsBar
        hasKnownWords={false}
        pauseCount={pauseCount}
        seekCount={seekCount}
        cardEventCount={cardEventCount}
        session={session}
        lookupRate={lookupRate}
      />
    </div>
  );
}

/* ── Stats Bar ──────────────────────────────────────────────────── */

function StatsBar({
  hasKnownWords,
  pauseCount,
  seekCount,
  cardEventCount,
  session,
  lookupRate,
}: {
  hasKnownWords: boolean;
  pauseCount: number;
  seekCount: number;
  cardEventCount: number;
  session: SessionSummary;
  lookupRate: ReturnType<typeof buildLookupRateDisplay>;
}) {
  return (
    <div className="flex flex-wrap items-center gap-4 text-[11px] pt-1">
      {/* Group 1: Legend */}
      {hasKnownWords && (
        <>
          <span className="flex items-center gap-1.5">
            <span
              className="inline-block w-2.5 h-2.5 rounded-sm"
              style={{ background: 'rgba(166,218,149,0.4)', border: '1px solid #a6da95' }}
            />
            <span className="text-ctp-overlay2">Known</span>
          </span>
          <span className="flex items-center gap-1.5">
            <span
              className="inline-block w-2.5 h-2.5 rounded-sm"
              style={{ background: 'rgba(198,160,246,0.2)', border: '1px solid #c6a0f6' }}
            />
            <span className="text-ctp-overlay2">Unknown</span>
          </span>
          <span className="text-ctp-surface2">|</span>
        </>
      )}

      {/* Group 2: Playback stats */}
      {pauseCount > 0 && (
        <span className="text-ctp-overlay2">
          <span className="text-ctp-peach">{pauseCount}</span> pause
          {pauseCount !== 1 ? 's' : ''}
        </span>
      )}
      {seekCount > 0 && (
        <span className="text-ctp-overlay2">
          <span className="text-ctp-teal">{seekCount}</span> seek{seekCount !== 1 ? 's' : ''}
        </span>
      )}
      {(pauseCount > 0 || seekCount > 0) && <span className="text-ctp-surface2">|</span>}

      {/* Group 3: Learning events */}
      <span className="flex items-center gap-1.5">
        <span
          className="inline-block w-3 h-0.5 rounded"
          style={{ background: '#b7bdf8', opacity: 0.8 }}
        />
        <span className="text-ctp-overlay2">
          {session.yomitanLookupCount} Yomitan lookup
          {session.yomitanLookupCount !== 1 ? 's' : ''}
        </span>
      </span>
      {lookupRate && (
        <span className="text-ctp-overlay2">
          lookup rate: <span className="text-ctp-sapphire">{lookupRate.shortValue}</span>{' '}
          <span className="text-ctp-subtext0">({lookupRate.longValue})</span>
        </span>
      )}
      <span className="flex items-center gap-1.5">
        <span className="text-[12px]">{'\u26CF'}</span>
        <span className="text-ctp-green">
          {Math.max(cardEventCount, session.cardsMined)} card
          {Math.max(cardEventCount, session.cardsMined) !== 1 ? 's' : ''} mined
        </span>
      </span>
    </div>
  );
}
