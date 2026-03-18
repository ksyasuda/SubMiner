import { formatEventSeconds, type SessionChartMarker, type SessionEventNoteInfo } from '../../lib/session-events';

interface SessionEventPopoverProps {
  marker: SessionChartMarker;
  noteInfos: Map<number, SessionEventNoteInfo>;
  loading: boolean;
  pinned: boolean;
  onTogglePinned: () => void;
  onClose: () => void;
  onOpenNote: (noteId: number) => void;
}

function formatEventTime(tsMs: number): string {
  return new Date(tsMs).toLocaleTimeString(undefined, {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

export function SessionEventPopover({
  marker,
  noteInfos,
  loading,
  pinned,
  onTogglePinned,
  onClose,
  onOpenNote,
}: SessionEventPopoverProps) {
  const seekDurationLabel =
    marker.kind === 'seek' && marker.fromMs !== null && marker.toMs !== null
      ? formatEventSeconds(Math.abs(marker.toMs - marker.fromMs))?.replace(/\.0s$/, 's')
      : null;

  return (
    <div className="relative z-50 w-64 rounded-xl border border-ctp-surface2 bg-ctp-surface0/95 p-3 shadow-2xl shadow-black/30 backdrop-blur-sm">
      <div className="mb-2 flex items-start justify-between gap-3">
        <div>
          <div className="text-xs font-semibold text-ctp-text">
            {marker.kind === 'pause' && 'Paused'}
            {marker.kind === 'seek' && `Seek ${marker.direction}`}
            {marker.kind === 'card' && 'Card mined'}
          </div>
          <div className="text-[10px] text-ctp-overlay1">{formatEventTime(marker.eventTsMs)}</div>
        </div>
        <div className="flex items-center gap-1.5">
          {pinned ? (
            <span className="rounded-full border border-ctp-surface2 px-1.5 py-0.5 text-[10px] font-medium text-ctp-blue">
              Pinned
            </span>
          ) : null}
          <button
            type="button"
            onClick={onTogglePinned}
            className="rounded-md border border-ctp-surface2 px-1.5 py-0.5 text-[10px] text-ctp-overlay1 transition-colors hover:bg-ctp-surface1 hover:text-ctp-text"
          >
            {pinned ? 'Unpin' : 'Pin'}
          </button>
          {pinned ? (
            <button
              type="button"
              aria-label="Close event popup"
              onClick={onClose}
              className="rounded-md border border-ctp-surface2 px-1.5 py-0.5 text-[10px] text-ctp-overlay1 transition-colors hover:bg-ctp-surface1 hover:text-ctp-text"
            >
              ×
            </button>
          ) : null}
          <div className="text-sm">
            {marker.kind === 'pause' && '||'}
            {marker.kind === 'seek' && (marker.direction === 'backward' ? '<<' : '>>')}
            {marker.kind === 'card' && '\u26CF'}
          </div>
        </div>
      </div>

      {marker.kind === 'pause' && (
        <div className="text-xs text-ctp-subtext0">
          Duration: <span className="text-ctp-peach">{formatEventSeconds(marker.durationMs)}</span>
        </div>
      )}

      {marker.kind === 'seek' && (
        <div className="space-y-1 text-xs text-ctp-subtext0">
          <div>
            From <span className="text-ctp-teal">{formatEventSeconds(marker.fromMs) ?? '\u2014'}</span>{' '}
            to <span className="text-ctp-teal">{formatEventSeconds(marker.toMs) ?? '\u2014'}</span>
          </div>
          <div>
            Length <span className="text-ctp-peach">{seekDurationLabel ?? '\u2014'}</span>
          </div>
        </div>
      )}

      {marker.kind === 'card' && (
        <div className="space-y-2">
          <div className="text-xs text-ctp-cards-mined">
            +{marker.cardsDelta} {marker.cardsDelta === 1 ? 'card' : 'cards'}
          </div>
          {loading ? (
            <div className="text-xs text-ctp-overlay1">Loading Anki note info...</div>
          ) : null}
          <div className="space-y-1.5">
            {marker.noteIds.length > 0 ? (
              marker.noteIds.map((noteId) => {
                const info = noteInfos.get(noteId);
                const hasPreview = Boolean(info?.expression || info?.context || info?.meaning);
                return (
                  <div
                    key={noteId}
                    className="rounded-lg border border-ctp-surface1 bg-ctp-mantle/80 px-2.5 py-2"
                  >
                    <div className="mb-1 flex items-center justify-between gap-2">
                      <div className="rounded-full bg-ctp-surface1 px-1.5 py-0.5 text-[10px] uppercase tracking-wide text-ctp-overlay1">
                        Note {noteId}
                      </div>
                      {!hasPreview ? (
                        <div className="text-[10px] text-ctp-overlay1">Preview unavailable</div>
                      ) : null}
                    </div>
                    {info?.expression ? (
                      <div className="mb-1 text-sm font-medium text-ctp-text">{info.expression}</div>
                    ) : null}
                    {info?.context ? (
                      <div className="mb-1 text-xs text-ctp-subtext0">{info.context}</div>
                    ) : null}
                    {info?.meaning ? (
                      <div className="mb-2 text-xs text-ctp-teal">{info.meaning}</div>
                    ) : null}
                    {!hasPreview ? (
                      <div className="mb-2 text-xs text-ctp-overlay1">
                        Preview unavailable from AnkiConnect.
                      </div>
                    ) : null}
                    <button
                      type="button"
                      onClick={() => onOpenNote(noteId)}
                      className="rounded-md bg-ctp-surface1 px-2 py-1 text-[10px] text-ctp-blue transition-colors hover:bg-ctp-surface2"
                    >
                      Open in Anki
                    </button>
                  </div>
                );
              })
            ) : (
              <div className="text-xs text-ctp-overlay1">No linked note ids recorded.</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
