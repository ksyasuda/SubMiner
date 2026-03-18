import { useEffect, useState } from 'react';
import { SessionDetail } from '../sessions/SessionDetail';
import { SessionRow } from '../sessions/SessionRow';
import type { SessionSummary } from '../../types/stats';

interface MediaSessionListProps {
  sessions: SessionSummary[];
  onDeleteSession: (session: SessionSummary) => void;
  deletingSessionId?: number | null;
  initialExpandedSessionId?: number | null;
  onConsumeInitialExpandedSession?: () => void;
}

export function MediaSessionList({
  sessions,
  onDeleteSession,
  deletingSessionId = null,
  initialExpandedSessionId = null,
  onConsumeInitialExpandedSession,
}: MediaSessionListProps) {
  const [expandedId, setExpandedId] = useState<number | null>(initialExpandedSessionId);

  useEffect(() => {
    if (initialExpandedSessionId == null) return;
    if (!sessions.some((session) => session.sessionId === initialExpandedSessionId)) return;
    setExpandedId(initialExpandedSessionId);
    onConsumeInitialExpandedSession?.();
  }, [initialExpandedSessionId, onConsumeInitialExpandedSession, sessions]);

  useEffect(() => {
    if (expandedId == null) return;
    if (sessions.some((session) => session.sessionId === expandedId)) return;
    setExpandedId(null);
  }, [expandedId, sessions]);

  if (sessions.length === 0) {
    return <div className="text-sm text-ctp-overlay2">No sessions recorded</div>;
  }

  return (
    <div className="space-y-2">
      <h3 className="text-sm font-semibold text-ctp-text">Session History</h3>
      {sessions.map((s) => (
        <div key={s.sessionId}>
          <SessionRow
            session={s}
            isExpanded={expandedId === s.sessionId}
            detailsId={`media-session-details-${s.sessionId}`}
            onToggle={() =>
              setExpandedId((current) => (current === s.sessionId ? null : s.sessionId))
            }
            onDelete={() => onDeleteSession(s)}
            deleteDisabled={deletingSessionId === s.sessionId}
          />
          {expandedId === s.sessionId ? (
            <div id={`media-session-details-${s.sessionId}`}>
              <SessionDetail session={s} />
            </div>
          ) : null}
        </div>
      ))}
    </div>
  );
}
