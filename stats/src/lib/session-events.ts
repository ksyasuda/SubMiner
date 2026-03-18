import { EventType, type SessionEvent } from '../types/stats';

export const SESSION_CHART_EVENT_TYPES = [
  EventType.CARD_MINED,
  EventType.SEEK_FORWARD,
  EventType.SEEK_BACKWARD,
  EventType.PAUSE_START,
  EventType.PAUSE_END,
  EventType.YOMITAN_LOOKUP,
] as const;

export interface PauseRegion {
  startMs: number;
  endMs: number;
}

export interface SessionChartEvents {
  cardEvents: SessionEvent[];
  seekEvents: SessionEvent[];
  yomitanLookupEvents: SessionEvent[];
  pauseRegions: PauseRegion[];
  markers: SessionChartMarker[];
}

export interface SessionEventNoteInfo {
  noteId: number;
  expression: string;
  context: string | null;
  meaning: string | null;
}

export interface SessionChartPlotArea {
  left: number;
  width: number;
}

interface SessionEventNoteField {
  value: string;
}

interface SessionEventNoteRecord {
  noteId: unknown;
  fields?: Record<string, SessionEventNoteField> | null;
}

export type SessionChartMarker =
  | {
      key: string;
      kind: 'pause';
      anchorTsMs: number;
      eventTsMs: number;
      startMs: number;
      endMs: number;
      durationMs: number;
    }
  | {
      key: string;
      kind: 'seek';
      anchorTsMs: number;
      eventTsMs: number;
      direction: 'forward' | 'backward';
      fromMs: number | null;
      toMs: number | null;
    }
  | {
      key: string;
      kind: 'card';
      anchorTsMs: number;
      eventTsMs: number;
      noteIds: number[];
      cardsDelta: number;
    };

function parsePayload(payload: string | null): Record<string, unknown> | null {
  if (!payload) return null;
  try {
    const parsed = JSON.parse(payload);
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : null;
  } catch {
    return null;
  }
}

function readNumberField(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function readNoteIds(value: unknown): number[] {
  if (!Array.isArray(value)) return [];
  return value.filter(
    (entry): entry is number => typeof entry === 'number' && Number.isInteger(entry),
  );
}

function stripHtml(value: string): string {
  return value
    .replace(/\[sound:[^\]]+\]/gi, ' ')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function pickFieldValue(
  fields: Record<string, SessionEventNoteField>,
  patterns: RegExp[],
  excludeValues: Set<string> = new Set(),
): string | null {
  const entries = Object.entries(fields);

  for (const pattern of patterns) {
    for (const [fieldName, field] of entries) {
      if (!pattern.test(fieldName)) continue;
      const cleaned = stripHtml(field?.value ?? '');
      if (cleaned && !excludeValues.has(cleaned)) return cleaned;
    }
  }

  return null;
}

function pickExpressionField(fields: Record<string, SessionEventNoteField>): string {
  const entries = Object.entries(fields);
  const preferredPatterns = [
    /^(expression|word|vocab|vocabulary|target|target word|front)$/i,
    /(expression|word|vocab|vocabulary|target)/i,
  ];

  const preferredValue = pickFieldValue(fields, preferredPatterns);
  if (preferredValue) return preferredValue;

  for (const [, field] of entries) {
    const cleaned = stripHtml(field?.value ?? '');
    if (cleaned) return cleaned;
  }

  return '';
}

export function extractSessionEventNoteInfo(
  note: SessionEventNoteRecord,
): SessionEventNoteInfo | null {
  if (typeof note.noteId !== 'number' || !Number.isInteger(note.noteId) || note.noteId <= 0) {
    return null;
  }

  const fields = note.fields ?? {};
  const expression = pickExpressionField(fields);
  const usedValues = new Set<string>(expression ? [expression] : []);
  const context =
    pickFieldValue(
      fields,
      [/^(sentence|context|example)$/i, /(sentence|context|example)/i],
      usedValues,
    ) ?? null;
  if (context) {
    usedValues.add(context);
  }
  const meaning =
    pickFieldValue(
      fields,
      [
        /^(meaning|definition|gloss|translation|back)$/i,
        /(meaning|definition|gloss|translation|back)/i,
      ],
      usedValues,
    ) ?? null;

  return {
    noteId: note.noteId,
    expression,
    context,
    meaning,
  };
}

export function resolveActiveSessionMarkerKey(
  hoveredMarkerKey: string | null,
  pinnedMarkerKey: string | null,
): string | null {
  return pinnedMarkerKey ?? hoveredMarkerKey;
}

export function togglePinnedSessionMarkerKey(
  currentPinnedMarkerKey: string | null,
  nextMarkerKey: string,
): string | null {
  return currentPinnedMarkerKey === nextMarkerKey ? null : nextMarkerKey;
}

export function formatEventSeconds(ms: number | null): string | null {
  if (ms == null || !Number.isFinite(ms)) return null;
  return `${(ms / 1000).toFixed(1)}s`;
}

export function projectSessionMarkerLeftPx({
  anchorTsMs,
  tsMin,
  tsMax,
  plotLeftPx,
  plotWidthPx,
}: {
  anchorTsMs: number;
  tsMin: number;
  tsMax: number;
  plotLeftPx: number;
  plotWidthPx: number;
}): number {
  if (plotWidthPx <= 0) return plotLeftPx;
  if (tsMax <= tsMin) return Math.round(plotLeftPx + plotWidthPx / 2);
  const ratio = Math.max(0, Math.min(1, (anchorTsMs - tsMin) / (tsMax - tsMin)));
  return Math.round(plotLeftPx + plotWidthPx * ratio);
}

export function buildSessionChartEvents(events: SessionEvent[]): SessionChartEvents {
  const cardEvents: SessionEvent[] = [];
  const seekEvents: SessionEvent[] = [];
  const yomitanLookupEvents: SessionEvent[] = [];
  const pauseRegions: PauseRegion[] = [];
  const markers: SessionChartMarker[] = [];
  let pendingPauseStartMs: number | null = null;

  for (const event of events) {
    switch (event.eventType) {
      case EventType.CARD_MINED:
        cardEvents.push(event);
        {
          const payload = parsePayload(event.payload);
          markers.push({
            key: `card-${event.tsMs}`,
            kind: 'card',
            anchorTsMs: event.tsMs,
            eventTsMs: event.tsMs,
            noteIds: readNoteIds(payload?.noteIds),
            cardsDelta: readNumberField(payload?.cardsMined) ?? 1,
          });
        }
        break;
      case EventType.SEEK_FORWARD:
      case EventType.SEEK_BACKWARD:
        seekEvents.push(event);
        {
          const payload = parsePayload(event.payload);
          markers.push({
            key: `seek-${event.tsMs}-${event.eventType}`,
            kind: 'seek',
            anchorTsMs: event.tsMs,
            eventTsMs: event.tsMs,
            direction: event.eventType === EventType.SEEK_BACKWARD ? 'backward' : 'forward',
            fromMs: readNumberField(payload?.fromMs),
            toMs: readNumberField(payload?.toMs),
          });
        }
        break;
      case EventType.YOMITAN_LOOKUP:
        yomitanLookupEvents.push(event);
        break;
      case EventType.PAUSE_START:
        pendingPauseStartMs = event.tsMs;
        break;
      case EventType.PAUSE_END:
        if (pendingPauseStartMs !== null) {
          pauseRegions.push({ startMs: pendingPauseStartMs, endMs: event.tsMs });
          markers.push({
            key: `pause-${pendingPauseStartMs}-${event.tsMs}`,
            kind: 'pause',
            anchorTsMs: pendingPauseStartMs + Math.round((event.tsMs - pendingPauseStartMs) / 2),
            eventTsMs: pendingPauseStartMs,
            startMs: pendingPauseStartMs,
            endMs: event.tsMs,
            durationMs: Math.max(0, event.tsMs - pendingPauseStartMs),
          });
          pendingPauseStartMs = null;
        }
        break;
      default:
        break;
    }
  }

  if (pendingPauseStartMs !== null) {
    pauseRegions.push({ startMs: pendingPauseStartMs, endMs: pendingPauseStartMs + 2_000 });
    markers.push({
      key: `pause-${pendingPauseStartMs}-${pendingPauseStartMs + 2_000}`,
      kind: 'pause',
      anchorTsMs: pendingPauseStartMs + 1_000,
      eventTsMs: pendingPauseStartMs,
      startMs: pendingPauseStartMs,
      endMs: pendingPauseStartMs + 2_000,
      durationMs: 2_000,
    });
  }

  markers.sort((left, right) => left.anchorTsMs - right.anchorTsMs);

  return {
    cardEvents,
    seekEvents,
    yomitanLookupEvents,
    pauseRegions,
    markers,
  };
}
