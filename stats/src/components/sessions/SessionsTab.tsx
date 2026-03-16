import { useEffect, useMemo, useState } from 'react';
import { useSessions } from '../../hooks/useSessions';
import { SessionRow } from './SessionRow';
import { SessionDetail } from './SessionDetail';
import { apiClient } from '../../lib/api-client';
import { confirmSessionDelete } from '../../lib/delete-confirm';
import { todayLocalDay, localDayFromMs } from '../../lib/formatters';
import type { SessionSummary } from '../../types/stats';

function groupSessionsByDay(sessions: SessionSummary[]): Map<string, SessionSummary[]> {
  const groups = new Map<string, SessionSummary[]>();
  const today = todayLocalDay();

  for (const session of sessions) {
    const sessionDay = localDayFromMs(session.startedAtMs);
    let label: string;
    if (sessionDay === today) {
      label = 'Today';
    } else if (sessionDay === today - 1) {
      label = 'Yesterday';
    } else {
      label = new Date(session.startedAtMs).toLocaleDateString(undefined, {
        month: 'long',
        day: 'numeric',
      });
    }
    const group = groups.get(label);
    if (group) {
      group.push(session);
    } else {
      groups.set(label, [session]);
    }
  }

  return groups;
}

export function SessionsTab() {
  const { sessions, loading, error } = useSessions();
  const [expandedId, setExpandedId] = useState<number | null>(null);
  const [search, setSearch] = useState('');
  const [visibleSessions, setVisibleSessions] = useState<SessionSummary[]>([]);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [deletingSessionId, setDeletingSessionId] = useState<number | null>(null);

  useEffect(() => {
    setVisibleSessions(sessions);
  }, [sessions]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return visibleSessions;
    return visibleSessions.filter((s) => s.canonicalTitle?.toLowerCase().includes(q));
  }, [visibleSessions, search]);

  const groups = useMemo(() => groupSessionsByDay(filtered), [filtered]);

  const handleDeleteSession = async (session: SessionSummary) => {
    if (!confirmSessionDelete()) return;

    setDeleteError(null);
    setDeletingSessionId(session.sessionId);
    try {
      await apiClient.deleteSession(session.sessionId);
      setVisibleSessions((prev) => prev.filter((item) => item.sessionId !== session.sessionId));
      setExpandedId((prev) => (prev === session.sessionId ? null : prev));
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete session.');
    } finally {
      setDeletingSessionId(null);
    }
  };

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;

  return (
    <div className="space-y-4">
      <input
        type="text"
        placeholder="Search by title..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2 text-sm text-ctp-text placeholder:text-ctp-overlay2 focus:outline-none focus:border-ctp-blue"
      />

      {deleteError ? <div className="text-sm text-ctp-red">{deleteError}</div> : null}

      {Array.from(groups.entries()).map(([dayLabel, daySessions]) => (
        <div key={dayLabel}>
          <div className="flex items-center gap-3 mb-2">
            <h3 className="text-xs font-semibold text-ctp-overlay2 uppercase tracking-widest shrink-0">
              {dayLabel}
            </h3>
            <div className="flex-1 h-px bg-gradient-to-r from-ctp-surface1 to-transparent" />
          </div>
          <div className="space-y-2">
            {daySessions.map((s) => {
              const detailsId = `session-details-${s.sessionId}`;
              return (
                <div key={s.sessionId}>
                  <SessionRow
                    session={s}
                    isExpanded={expandedId === s.sessionId}
                    detailsId={detailsId}
                    onToggle={() => setExpandedId(expandedId === s.sessionId ? null : s.sessionId)}
                    onDelete={() => void handleDeleteSession(s)}
                    deleteDisabled={deletingSessionId === s.sessionId}
                  />
                  {expandedId === s.sessionId && (
                    <div id={detailsId}>
                      <SessionDetail sessionId={s.sessionId} cardsMined={s.cardsMined} />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      ))}

      {filtered.length === 0 && (
        <div className="text-ctp-overlay2 text-sm">
          {search.trim() ? 'No sessions matching your search.' : 'No sessions recorded yet.'}
        </div>
      )}
    </div>
  );
}
