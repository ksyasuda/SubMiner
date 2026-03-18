import { useState, useEffect } from 'react';
import { getStatsClient } from '../../hooks/useStatsApi';
import { apiClient } from '../../lib/api-client';
import { confirmSessionDelete } from '../../lib/delete-confirm';
import { formatDuration, formatNumber, formatRelativeDate } from '../../lib/formatters';
import { getSessionDisplayWordCount } from '../../lib/session-word-count';
import type { EpisodeDetailData } from '../../types/stats';

interface EpisodeDetailProps {
  videoId: number;
  onSessionDeleted?: () => void;
}

interface NoteInfo {
  noteId: number;
  expression: string;
}

export function EpisodeDetail({ videoId, onSessionDeleted }: EpisodeDetailProps) {
  const [data, setData] = useState<EpisodeDetailData | null>(null);
  const [loading, setLoading] = useState(true);
  const [noteInfos, setNoteInfos] = useState<Map<number, NoteInfo>>(new Map());

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    getStatsClient()
      .getEpisodeDetail(videoId)
      .then((d) => {
        if (cancelled) return;
        setData(d);
        const allNoteIds = d.cardEvents.flatMap((ev) => ev.noteIds);
        if (allNoteIds.length > 0) {
          getStatsClient()
            .ankiNotesInfo(allNoteIds)
            .then((notes) => {
              if (cancelled) return;
              const map = new Map<number, NoteInfo>();
              for (const note of notes) {
                const expr = note.preview?.word ?? '';
                map.set(note.noteId, { noteId: note.noteId, expression: expr });
              }
              setNoteInfos(map);
            })
            .catch((err) => console.warn('Failed to fetch Anki note info:', err));
        }
      })
      .catch(() => {
        if (!cancelled) setData(null);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [videoId]);

  const handleDeleteSession = async (sessionId: number) => {
    if (!confirmSessionDelete()) return;
    await apiClient.deleteSession(sessionId);
    setData((prev) => {
      if (!prev) return prev;
      return { ...prev, sessions: prev.sessions.filter((s) => s.sessionId !== sessionId) };
    });
    onSessionDeleted?.();
  };

  if (loading) return <div className="text-ctp-overlay2 text-xs p-3">Loading...</div>;
  if (!data)
    return <div className="text-ctp-overlay2 text-xs p-3">Failed to load episode details.</div>;

  const { sessions, cardEvents } = data;

  return (
    <div className="bg-ctp-mantle border border-ctp-surface1 rounded-lg">
      {sessions.length > 0 && (
        <div className="p-3 border-b border-ctp-surface1">
          <h4 className="text-xs font-semibold text-ctp-subtext0 mb-2">Sessions</h4>
          <div className="space-y-1">
            {sessions.map((s) => (
              <div key={s.sessionId} className="flex items-center gap-3 text-xs group">
                <span className="text-ctp-overlay2">
                  {s.startedAtMs > 0 ? formatRelativeDate(s.startedAtMs) : '\u2014'}
                </span>
                <span className="text-ctp-blue">{formatDuration(s.activeWatchedMs)}</span>
                <span className="text-ctp-cards-mined">{formatNumber(s.cardsMined)} cards</span>
                <span className="text-ctp-peach">
                  {formatNumber(getSessionDisplayWordCount(s))} tokens
                </span>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    void handleDeleteSession(s.sessionId);
                  }}
                  className="ml-auto opacity-0 group-hover:opacity-100 text-ctp-red/70 hover:text-ctp-red transition-opacity text-[10px] px-1.5 py-0.5 rounded hover:bg-ctp-red/10"
                  title="Delete session"
                >
                  Delete
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {cardEvents.length > 0 && (
        <div className="p-3 border-b border-ctp-surface1">
          <h4 className="text-xs font-semibold text-ctp-subtext0 mb-2">Cards Mined</h4>
          <div className="space-y-1.5">
            {cardEvents.map((ev) => (
              <div key={ev.eventId} className="flex items-center gap-2 text-xs">
                <span className="text-ctp-overlay2 shrink-0">{formatRelativeDate(ev.tsMs)}</span>
                {ev.noteIds.length > 0 ? (
                  ev.noteIds.map((noteId) => {
                    const info = noteInfos.get(noteId);
                    return (
                      <div key={noteId} className="flex items-center gap-2 min-w-0 flex-1">
                        {info?.expression && (
                          <span className="text-ctp-text font-medium truncate">
                            {info.expression}
                          </span>
                        )}
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            getStatsClient().ankiBrowse(noteId);
                          }}
                          className="px-1.5 py-0.5 bg-ctp-surface1 text-ctp-blue rounded text-[10px] hover:bg-ctp-surface2 transition-colors cursor-pointer shrink-0 ml-auto"
                        >
                          Open in Anki
                        </button>
                      </div>
                    );
                  })
                ) : (
                  <span className="text-ctp-cards-mined">
                    +{ev.cardsDelta} {ev.cardsDelta === 1 ? 'card' : 'cards'}
                  </span>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {sessions.length === 0 && cardEvents.length === 0 && (
        <div className="p-3 text-xs text-ctp-overlay2">No detailed data available.</div>
      )}
    </div>
  );
}
