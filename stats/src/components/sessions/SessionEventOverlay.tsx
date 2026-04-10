import { useEffect, useRef, type FocusEvent, type MouseEvent } from 'react';
import {
  projectSessionMarkerLeftPx,
  resolveActiveSessionMarkerKey,
  togglePinnedSessionMarkerKey,
  type SessionChartMarker,
  type SessionEventNoteInfo,
  type SessionChartPlotArea,
} from '../../lib/session-events';
import { SessionEventPopover } from './SessionEventPopover';

interface SessionEventOverlayProps {
  markers: SessionChartMarker[];
  tsMin: number;
  tsMax: number;
  plotArea: SessionChartPlotArea | null;
  hoveredMarkerKey: string | null;
  onHoveredMarkerChange: (markerKey: string | null) => void;
  pinnedMarkerKey: string | null;
  onPinnedMarkerChange: (markerKey: string | null) => void;
  noteInfos: Map<number, SessionEventNoteInfo>;
  loadingNoteIds: Set<number>;
  onOpenNote: (noteId: number) => void;
}

function toPercent(tsMs: number, tsMin: number, tsMax: number): number {
  if (tsMax <= tsMin) return 50;
  const ratio = ((tsMs - tsMin) / (tsMax - tsMin)) * 100;
  return Math.max(0, Math.min(100, ratio));
}

function markerLabel(marker: SessionChartMarker): string {
  switch (marker.kind) {
    case 'pause':
      return '||';
    case 'card':
      return '\u26CF';
  }
}

function markerColors(marker: SessionChartMarker): { border: string; bg: string; text: string } {
  switch (marker.kind) {
    case 'pause':
      return { border: '#f5a97f', bg: 'rgba(245,169,127,0.16)', text: '#f5a97f' };
    case 'card':
      return { border: '#a6da95', bg: 'rgba(166,218,149,0.16)', text: '#a6da95' };
  }
}

function popupAlignment(percent: number): string {
  if (percent <= 15) return 'left-0 translate-x-0';
  if (percent >= 85) return 'right-0 translate-x-0';
  return 'left-1/2 -translate-x-1/2';
}

function handleWrapperBlur(
  event: FocusEvent<HTMLDivElement>,
  onHoveredMarkerChange: (markerKey: string | null) => void,
  pinnedMarkerKey: string | null,
  markerKey: string,
): void {
  if (pinnedMarkerKey === markerKey) return;
  const nextFocused = event.relatedTarget;
  if (nextFocused instanceof Node && event.currentTarget.contains(nextFocused)) {
    return;
  }
  onHoveredMarkerChange(null);
}

function handleWrapperMouseLeave(
  event: MouseEvent<HTMLDivElement>,
  onHoveredMarkerChange: (markerKey: string | null) => void,
  pinnedMarkerKey: string | null,
  markerKey: string,
): void {
  if (pinnedMarkerKey === markerKey) return;
  const nextHovered = event.relatedTarget;
  if (nextHovered instanceof Node && event.currentTarget.contains(nextHovered)) {
    return;
  }
  onHoveredMarkerChange(null);
}

export function SessionEventOverlay({
  markers,
  tsMin,
  tsMax,
  plotArea,
  hoveredMarkerKey,
  onHoveredMarkerChange,
  pinnedMarkerKey,
  onPinnedMarkerChange,
  noteInfos,
  loadingNoteIds,
  onOpenNote,
}: SessionEventOverlayProps) {
  if (markers.length === 0) return null;

  const rootRef = useRef<HTMLDivElement>(null);
  const activeMarkerKey = resolveActiveSessionMarkerKey(hoveredMarkerKey, pinnedMarkerKey);

  useEffect(() => {
    if (!pinnedMarkerKey) return;

    function handleDocumentPointerDown(event: PointerEvent): void {
      if (rootRef.current?.contains(event.target as Node)) {
        return;
      }
      onPinnedMarkerChange(null);
      onHoveredMarkerChange(null);
    }

    function handleDocumentKeyDown(event: KeyboardEvent): void {
      if (event.key !== 'Escape') return;
      onPinnedMarkerChange(null);
      onHoveredMarkerChange(null);
    }

    document.addEventListener('pointerdown', handleDocumentPointerDown);
    document.addEventListener('keydown', handleDocumentKeyDown);
    return () => {
      document.removeEventListener('pointerdown', handleDocumentPointerDown);
      document.removeEventListener('keydown', handleDocumentKeyDown);
    };
  }, [pinnedMarkerKey, onHoveredMarkerChange, onPinnedMarkerChange]);

  return (
    <div ref={rootRef} className="pointer-events-none absolute inset-0 z-30 overflow-visible">
      {markers.map((marker) => {
        const percent = toPercent(marker.anchorTsMs, tsMin, tsMax);
        const left = plotArea
          ? `${projectSessionMarkerLeftPx({
              anchorTsMs: marker.anchorTsMs,
              tsMin,
              tsMax,
              plotLeftPx: plotArea.left,
              plotWidthPx: plotArea.width,
            })}px`
          : `${percent}%`;
        const colors = markerColors(marker);
        const isActive = marker.key === activeMarkerKey;
        const isPinned = marker.key === pinnedMarkerKey;
        const loading =
          marker.kind === 'card' && marker.noteIds.some((noteId) => loadingNoteIds.has(noteId));

        return (
          <div
            key={marker.key}
            className="pointer-events-auto absolute top-0 -translate-x-1/2 pt-1"
            style={{ left }}
            onMouseEnter={() => onHoveredMarkerChange(marker.key)}
            onMouseLeave={(event) =>
              handleWrapperMouseLeave(event, onHoveredMarkerChange, pinnedMarkerKey, marker.key)
            }
            onFocusCapture={() => onHoveredMarkerChange(marker.key)}
            onBlurCapture={(event) =>
              handleWrapperBlur(event, onHoveredMarkerChange, pinnedMarkerKey, marker.key)
            }
          >
            <div className="relative flex flex-col items-center">
              <button
                type="button"
                aria-label={`Show ${marker.kind} event details`}
                aria-pressed={isPinned}
                className="flex h-5 min-w-5 items-center justify-center rounded-full border px-1 text-[10px] font-semibold shadow-sm backdrop-blur-sm"
                style={{
                  borderColor: colors.border,
                  background: colors.bg,
                  color: colors.text,
                }}
                onClick={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  onHoveredMarkerChange(marker.key);
                  onPinnedMarkerChange(togglePinnedSessionMarkerKey(pinnedMarkerKey, marker.key));
                }}
              >
                {markerLabel(marker)}
              </button>
              {isActive ? (
                <div
                  className={`pointer-events-auto absolute top-5 z-50 pt-2 ${popupAlignment(percent)}`}
                  onMouseDownCapture={() => {
                    if (!isPinned) {
                      onPinnedMarkerChange(marker.key);
                    }
                  }}
                >
                  <SessionEventPopover
                    marker={marker}
                    noteInfos={noteInfos}
                    loading={loading}
                    pinned={isPinned}
                    onTogglePinned={() =>
                      onPinnedMarkerChange(
                        togglePinnedSessionMarkerKey(pinnedMarkerKey, marker.key),
                      )
                    }
                    onClose={() => {
                      onPinnedMarkerChange(null);
                      onHoveredMarkerChange(null);
                    }}
                    onOpenNote={onOpenNote}
                  />
                </div>
              ) : null}
            </div>
          </div>
        );
      })}
    </div>
  );
}
