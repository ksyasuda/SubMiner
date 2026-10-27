import { useEffect, useMemo, useRef, useState } from 'react';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Customized,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { useSessionDetail } from '../../hooks/useSessions';
import { getStatsClient } from '../../hooks/useStatsApi';
import { CHART_THEME } from '../../lib/chart-theme';
import { formatClockDuration, formatDuration, formatNumber } from '../../lib/formatters';
import {
  buildSessionDensity,
  formatBucketSize,
  type SessionDensity,
  type SessionDensityBucket,
} from '../../lib/session-density';
import {
  buildSessionChartEvents,
  collectPendingSessionEventNoteIds,
  getSessionEventCardRequest,
  mergeSessionEventNoteInfos,
  projectSessionMarkerLeftPx,
  resolveActiveSessionMarkerKey,
  type SessionChartMarker,
  type SessionEventNoteInfo,
  type SessionChartPlotArea,
} from '../../lib/session-events';
import { getSessionDisplayWordCount } from '../../lib/session-word-count';
import { buildLookupRateDisplay } from '../../lib/yomitan-lookup';
import type { SessionEvent, SessionSummary } from '../../types/stats';
import { SessionEventOverlay } from './SessionEventOverlay';

interface SessionDetailProps {
  session: SessionSummary;
}

const SERIES_COLORS = {
  known: '#a6da95',
  unknown: '#c6a0f6',
  words: '#8aadf4',
  lookup: '#b7bdf8',
  card: '#f5bde6',
} as const;

const CHART_HEIGHT = 150;

function formatTime(ms: number): string {
  return new Date(ms).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
}

function formatTimeWithSeconds(ms: number): string {
  return new Date(ms).toLocaleTimeString(undefined, {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

/** Reads the plot offset recharts computed so the event lane can line up with the bars. */
function SessionChartOffsetProbe({
  offset,
  onPlotAreaChange,
}: {
  offset?: { left?: number; width?: number };
  onPlotAreaChange: (plotArea: SessionChartPlotArea) => void;
}) {
  useEffect(() => {
    if (!offset) return;
    const { left, width } = offset;
    if (typeof left !== 'number' || !Number.isFinite(left)) return;
    if (typeof width !== 'number' || !Number.isFinite(width)) return;
    onPlotAreaChange({ left, width });
  }, [offset?.left, offset?.width, onPlotAreaChange]);

  return null;
}

export function SessionDetail({ session }: SessionDetailProps) {
  const { timeline, events, knownWordsTimeline, loading, error } = useSessionDetail(
    session.sessionId,
  );
  const [hoveredMarkerKey, setHoveredMarkerKey] = useState<string | null>(null);
  const [pinnedMarkerKey, setPinnedMarkerKey] = useState<string | null>(null);
  const [noteInfos, setNoteInfos] = useState<Map<number, SessionEventNoteInfo>>(new Map());
  const [loadingNoteIds, setLoadingNoteIds] = useState<Set<number>>(new Set());
  const pendingNoteIdsRef = useRef<Set<number>>(new Set());

  const density = useMemo(
    () => buildSessionDensity(timeline, knownWordsTimeline, events),
    [timeline, knownWordsTimeline, events],
  );
  const { yomitanLookupEvents, markers } = useMemo(() => buildSessionChartEvents(events), [events]);
  const activeMarkerKey = resolveActiveSessionMarkerKey(hoveredMarkerKey, pinnedMarkerKey);
  const activeMarker = useMemo<SessionChartMarker | null>(
    () => markers.find((marker) => marker.key === activeMarkerKey) ?? null,
    [markers, activeMarkerKey],
  );
  const activeCardRequest = useMemo(
    () => getSessionEventCardRequest(activeMarker),
    [activeMarkerKey, markers],
  );

  useEffect(() => {
    if (!activeCardRequest.requestKey || activeCardRequest.noteIds.length === 0) {
      return;
    }

    const missingNoteIds = collectPendingSessionEventNoteIds(
      activeCardRequest.noteIds,
      noteInfos,
      pendingNoteIdsRef.current,
    );
    if (missingNoteIds.length === 0) {
      return;
    }

    for (const noteId of missingNoteIds) {
      pendingNoteIdsRef.current.add(noteId);
    }

    let cancelled = false;
    setLoadingNoteIds((prev) => {
      const next = new Set(prev);
      for (const noteId of missingNoteIds) {
        next.add(noteId);
      }
      return next;
    });

    getStatsClient()
      .ankiNotesInfo(missingNoteIds)
      .then((notes) => {
        if (cancelled) return;
        setNoteInfos((prev) => {
          const next = new Map(prev);
          for (const [noteId, info] of mergeSessionEventNoteInfos(missingNoteIds, notes)) {
            next.set(noteId, info);
          }
          return next;
        });
      })
      .catch((err) => {
        if (!cancelled) {
          console.warn('Failed to fetch session event Anki note info:', err);
        }
      })
      .finally(() => {
        if (cancelled) return;
        for (const noteId of missingNoteIds) {
          pendingNoteIdsRef.current.delete(noteId);
        }
        setLoadingNoteIds((prev) => {
          const next = new Set(prev);
          for (const noteId of missingNoteIds) {
            next.delete(noteId);
          }
          return next;
        });
      });

    return () => {
      cancelled = true;
      for (const noteId of missingNoteIds) {
        pendingNoteIdsRef.current.delete(noteId);
      }
      setLoadingNoteIds((prev) => {
        const next = new Set(prev);
        for (const noteId of missingNoteIds) {
          next.delete(noteId);
        }
        return next;
      });
    };
  }, [activeCardRequest.requestKey, noteInfos]);

  const handleOpenNote = (noteId: number) => {
    void getStatsClient().ankiBrowse(noteId);
  };

  if (loading) return <div className="text-ctp-overlay2 text-xs p-2">Loading session...</div>;
  if (error) return <div className="text-ctp-red text-xs p-2">Error: {error}</div>;

  return (
    <div className="bg-ctp-mantle border border-ctp-surface1 rounded-lg p-3 mt-1 space-y-3">
      <SessionStatTiles session={session} />
      {density.buckets.length === 0 ? (
        <div className="text-ctp-overlay2 text-xs">No activity samples for this session.</div>
      ) : (
        <SessionDensityChart
          density={density}
          lookupEvents={yomitanLookupEvents}
          markers={markers}
          hoveredMarkerKey={hoveredMarkerKey}
          onHoveredMarkerChange={setHoveredMarkerKey}
          pinnedMarkerKey={pinnedMarkerKey}
          onPinnedMarkerChange={setPinnedMarkerKey}
          noteInfos={noteInfos}
          loadingNoteIds={loadingNoteIds}
          onOpenNote={handleOpenNote}
        />
      )}
    </div>
  );
}

/* ── Stat tiles ─────────────────────────────────────────────────── */

export function buildSessionStatTiles(session: SessionSummary): Array<{
  key: string;
  label: string;
  value: string;
  detail: string;
  colorClass: string;
}> {
  const words = getSessionDisplayWordCount(session);
  const activeMin = session.activeWatchedMs / 60_000;
  const wordsPerMin = activeMin > 0 ? words / activeMin : 0;
  const linesPerMin = activeMin > 0 ? session.linesSeen / activeMin : 0;
  const pausedMs = Math.max(0, session.totalWatchedMs - session.activeWatchedMs);
  const pausedPct =
    session.totalWatchedMs > 0 ? Math.round((pausedMs / session.totalWatchedMs) * 100) : 0;
  const lookupRate = buildLookupRateDisplay(session.yomitanLookupCount, words);
  const hasKnownRate = session.knownWordsSeen > 0 && words > 0;

  return [
    {
      key: 'pace',
      label: 'Pace',
      value: activeMin > 0 ? `${wordsPerMin.toFixed(0)} words/min` : '—',
      detail: activeMin > 0 ? `${linesPerMin.toFixed(1)} lines/min` : 'no active time',
      colorClass: 'text-ctp-blue',
    },
    {
      key: 'lines',
      label: 'Lines',
      value: formatNumber(session.linesSeen),
      detail: `${formatNumber(words)} words`,
      colorClass: 'text-ctp-mauve',
    },
    {
      key: 'known',
      label: 'Known',
      value: hasKnownRate ? `${session.knownWordRate.toFixed(0)}%` : '—',
      detail: hasKnownRate
        ? `${formatNumber(session.knownWordsSeen)} known words`
        : 'no known-word data',
      colorClass: 'text-ctp-green',
    },
    {
      key: 'lookups',
      label: 'Lookups',
      value: formatNumber(session.yomitanLookupCount),
      detail: lookupRate ? lookupRate.longValue : 'Yomitan lookups',
      colorClass: 'text-ctp-lavender',
    },
    {
      key: 'paused',
      label: 'Paused',
      value: pausedMs > 0 ? formatClockDuration(pausedMs) : 'none',
      detail: `${pausedPct}% of ${formatDuration(session.totalWatchedMs)} wall time`,
      colorClass: 'text-ctp-peach',
    },
  ];
}

function SessionStatTiles({ session }: { session: SessionSummary }) {
  const tiles = buildSessionStatTiles(session);
  return (
    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2">
      {tiles.map((tile) => (
        <div
          key={tile.key}
          className="rounded-md border border-ctp-surface1 bg-ctp-base/60 px-2.5 py-2 min-w-0"
        >
          <div className="text-[10px] uppercase tracking-wide text-ctp-overlay1">{tile.label}</div>
          <div className={`text-sm font-semibold tabular-nums truncate ${tile.colorClass}`}>
            {tile.value}
          </div>
          <div className="text-[10px] text-ctp-overlay2 truncate">{tile.detail}</div>
        </div>
      ))}
    </div>
  );
}

/* ── Density chart + event lane ─────────────────────────────────── */

interface SessionDensityChartProps {
  density: SessionDensity;
  lookupEvents: SessionEvent[];
  markers: SessionChartMarker[];
  hoveredMarkerKey: string | null;
  onHoveredMarkerChange: (markerKey: string | null) => void;
  pinnedMarkerKey: string | null;
  onPinnedMarkerChange: (markerKey: string | null) => void;
  noteInfos: Map<number, SessionEventNoteInfo>;
  loadingNoteIds: Set<number>;
  onOpenNote: (noteId: number) => void;
}

/** Roughly five evenly spaced tick indexes, always including the first bucket. */
export function pickDensityTickIndexes(bucketCount: number, maxTicks = 5): number[] {
  if (bucketCount <= 0) return [];
  const step = Math.max(1, Math.ceil(bucketCount / maxTicks));
  const ticks: number[] = [];
  for (let index = 0; index < bucketCount; index += step) {
    ticks.push(index);
  }
  return ticks;
}

function DensityTooltip({
  active,
  payload,
  hasKnownWords,
}: {
  active?: boolean;
  payload?: Array<{ payload?: SessionDensityBucket }>;
  hasKnownWords: boolean;
}) {
  const bucket = payload?.[0]?.payload;
  if (!active || !bucket) return null;
  const knownPct = bucket.words > 0 ? Math.round((bucket.knownWords / bucket.words) * 100) : 0;

  return (
    <div className="rounded-md border border-ctp-surface1 bg-ctp-surface0 px-2.5 py-2 text-[11px] text-ctp-text shadow-lg">
      <div className="text-ctp-subtext0 tabular-nums">
        {formatTimeWithSeconds(bucket.startMs)} – {formatTimeWithSeconds(bucket.endMs)}
      </div>
      <div className="mt-1 font-semibold tabular-nums">
        {formatNumber(bucket.words)} {bucket.words === 1 ? 'word' : 'words'}
      </div>
      {hasKnownWords && bucket.words > 0 ? (
        <div className="tabular-nums">
          <span style={{ color: SERIES_COLORS.known }}>
            {formatNumber(bucket.knownWords)} known
          </span>
          <span className="text-ctp-overlay1"> ({knownPct}%)</span>
          <span className="text-ctp-overlay1"> · </span>
          <span style={{ color: SERIES_COLORS.unknown }}>
            {formatNumber(bucket.unknownWords)} unknown
          </span>
        </div>
      ) : null}
      {bucket.lookups > 0 ? (
        <div className="tabular-nums" style={{ color: SERIES_COLORS.lookup }}>
          {bucket.lookups} {bucket.lookups === 1 ? 'lookup' : 'lookups'}
        </div>
      ) : null}
      {bucket.cards > 0 ? (
        <div className="tabular-nums" style={{ color: SERIES_COLORS.card }}>
          {bucket.cards} {bucket.cards === 1 ? 'card' : 'cards'} mined
        </div>
      ) : null}
    </div>
  );
}

function SessionDensityChart({
  density,
  lookupEvents,
  markers,
  hoveredMarkerKey,
  onHoveredMarkerChange,
  pinnedMarkerKey,
  onPinnedMarkerChange,
  noteInfos,
  loadingNoteIds,
  onOpenNote,
}: SessionDensityChartProps) {
  const [plotArea, setPlotArea] = useState<SessionChartPlotArea | null>(null);
  const { buckets, startMs, endMs, hasKnownWords } = density;
  const tickIndexes = pickDensityTickIndexes(buckets.length);
  const hasLane = lookupEvents.length > 0 || markers.length > 0;

  const laneLeft = (tsMs: number): string => {
    if (plotArea) {
      return `${projectSessionMarkerLeftPx({
        anchorTsMs: tsMs,
        tsMin: startMs,
        tsMax: endMs,
        plotLeftPx: plotArea.left,
        plotWidthPx: plotArea.width,
      })}px`;
    }
    const ratio = endMs > startMs ? (tsMs - startMs) / (endMs - startMs) : 0.5;
    return `${Math.max(0, Math.min(100, ratio * 100))}%`;
  };

  return (
    <div>
      <div className="flex items-baseline justify-between mb-1">
        <div className="text-xs font-semibold text-ctp-text">Reading density</div>
        <div className="text-[10px] text-ctp-overlay1">
          words per {formatBucketSize(density.bucketMs)}
        </div>
      </div>

      <ResponsiveContainer width="100%" height={CHART_HEIGHT}>
        <BarChart
          data={buckets}
          margin={{ top: 4, right: 4, bottom: 0, left: 0 }}
          barCategoryGap="22%"
        >
          <Customized
            component={
              <SessionChartOffsetProbe
                onPlotAreaChange={(nextPlotArea) => {
                  setPlotArea((prevPlotArea) =>
                    prevPlotArea &&
                    prevPlotArea.left === nextPlotArea.left &&
                    prevPlotArea.width === nextPlotArea.width
                      ? prevPlotArea
                      : nextPlotArea,
                  );
                }}
              />
            }
          />
          <CartesianGrid
            horizontal
            vertical={false}
            stroke={CHART_THEME.grid}
            strokeDasharray="3 3"
            strokeOpacity={0.35}
          />
          <XAxis
            dataKey="index"
            type="category"
            ticks={tickIndexes}
            interval={0}
            tickFormatter={(index: number) => formatTime(buckets[index]?.startMs ?? startMs)}
            tick={{ fontSize: 9, fill: CHART_THEME.tick }}
            axisLine={false}
            tickLine={false}
          />
          <YAxis
            tick={{ fontSize: 9, fill: CHART_THEME.tick }}
            axisLine={false}
            tickLine={false}
            width={30}
            allowDecimals={false}
          />
          <Tooltip
            cursor={{ fill: 'rgba(202,211,245,0.08)' }}
            content={<DensityTooltip hasKnownWords={hasKnownWords} />}
            isAnimationActive={false}
          />
          {/* recharts only registers direct Bar children, so no fragment wrapper here */}
          {hasKnownWords ? (
            <Bar
              key="known"
              dataKey="knownWords"
              stackId="words"
              name="Known"
              fill={SERIES_COLORS.known}
              fillOpacity={0.85}
              isAnimationActive={false}
            />
          ) : null}
          {hasKnownWords ? (
            <Bar
              key="unknown"
              dataKey="unknownWords"
              stackId="words"
              name="Unknown"
              fill={SERIES_COLORS.unknown}
              fillOpacity={0.85}
              radius={[3, 3, 0, 0]}
              isAnimationActive={false}
            />
          ) : (
            <Bar
              key="words"
              dataKey="words"
              name="Words"
              fill={SERIES_COLORS.words}
              fillOpacity={0.85}
              radius={[3, 3, 0, 0]}
              isAnimationActive={false}
            />
          )}
        </BarChart>
      </ResponsiveContainer>

      {hasLane ? (
        <div className="relative h-7" aria-label="Lookup and card events">
          <div
            className="absolute top-1/2 h-px bg-ctp-surface1"
            style={
              plotArea ? { left: plotArea.left, width: plotArea.width } : { left: 0, right: 0 }
            }
          />
          {lookupEvents.map((event, index) => (
            <span
              key={`lookup-${event.tsMs}-${index}`}
              title={`Yomitan lookup at ${formatTimeWithSeconds(event.tsMs)}`}
              className="absolute top-1.5 bottom-1.5 w-0.5 -translate-x-1/2 rounded-full"
              style={{ left: laneLeft(event.tsMs), background: SERIES_COLORS.lookup, opacity: 0.8 }}
            />
          ))}
          <SessionEventOverlay
            markers={markers}
            tsMin={startMs}
            tsMax={endMs}
            plotArea={plotArea}
            hoveredMarkerKey={hoveredMarkerKey}
            onHoveredMarkerChange={onHoveredMarkerChange}
            pinnedMarkerKey={pinnedMarkerKey}
            onPinnedMarkerChange={onPinnedMarkerChange}
            noteInfos={noteInfos}
            loadingNoteIds={loadingNoteIds}
            onOpenNote={onOpenNote}
          />
        </div>
      ) : null}

      <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[10px] text-ctp-overlay2 mt-1">
        {hasKnownWords ? (
          <>
            <LegendSwatch color={SERIES_COLORS.known} label="Known words" />
            <LegendSwatch color={SERIES_COLORS.unknown} label="Unknown words" />
          </>
        ) : (
          <LegendSwatch color={SERIES_COLORS.words} label="Words" />
        )}
        {lookupEvents.length > 0 ? (
          <span className="flex items-center gap-1.5">
            <span
              className="inline-block w-0.5 h-3 rounded-full"
              style={{ background: SERIES_COLORS.lookup }}
            />
            Yomitan lookup
          </span>
        ) : null}
        {markers.length > 0 ? (
          <span className="flex items-center gap-1.5">
            <span style={{ color: SERIES_COLORS.card }}>{'⛏'}</span>
            Card mined (click for note)
          </span>
        ) : null}
      </div>
    </div>
  );
}

function LegendSwatch({ color, label }: { color: string; label: string }) {
  return (
    <span className="flex items-center gap-1.5">
      <span className="inline-block w-2.5 h-2.5 rounded-sm" style={{ background: color }} />
      {label}
    </span>
  );
}
