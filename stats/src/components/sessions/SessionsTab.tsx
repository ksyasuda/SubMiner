import { useEffect, useMemo, useState } from 'react';
import { useSessions } from '../../hooks/useSessions';
import { SessionRow } from './SessionRow';
import { SessionDetail } from './SessionDetail';
import { apiClient } from '../../lib/api-client';
import { confirmSessionDelete } from '../../lib/delete-confirm';
import { formatSessionDayLabel } from '../../lib/formatters';
import type { SessionSummary } from '../../types/stats';

function groupSessionsByDay(sessions: SessionSummary[]): Map<string, SessionSummary[]> {
  const groups = new Map<string, SessionSummary[]>();

  for (const session of sessions) {
    const dayLabel = formatSessionDayLabel(session.startedAtMs);
    const group = groups.get(dayLabel);
    if (group) {
      group.push(session);
    } else {
      groups.set(dayLabel, [session]);
    }
  }

  return groups;
}

interface SessionsTabProps {
  initialSessionId?: number | null;
  onClearInitialSession?: () => void;
  onNavigateToMediaDetail?: (videoId: number) => void;
}

export function SessionsTab({
  initialSessionId,
  onClearInitialSession,
  onNavigateToMediaDetail,
}: SessionsTabProps = {}) {
  const { sessions, loading, error } = useSessions();
  const [expandedId, setExpandedId] = useState<number | null>(null);
  const [search, setSearch] = useState('');
  const [visibleSessions, setVisibleSessions] = useState<SessionSummary[]>([]);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [deletingSessionId, setDeletingSessionId] = useState<number | null>(null);

  useEffect(() => {
    setVisibleSessions(sessions);
  }, [sessions]);

  useEffect(() => {
    if (initialSessionId != null && sessions.length > 0) {
      let canceled = false;
      setExpandedId(initialSessionId);
      onClearInitialSession?.();
      const frame = requestAnimationFrame(() => {
        if (canceled) return;
        const el = document.getElementById(`session-details-${initialSessionId}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } else {
          // Session row itself if detail hasn't rendered yet
          const row = document.querySelector(
            `[aria-controls="session-details-${initialSessionId}"]`,
          );
          row?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      });
      return () => {
        canceled = true;
        cancelAnimationFrame(frame);
      };
    }
  }, [initialSessionId, sessions, onClearInitialSession]);

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
        type="search"
        aria-label="Search sessions by title"
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
                    onNavigateToMediaDetail={onNavigateToMediaDetail}
                  />
                  {expandedId === s.sessionId && (
                    <div id={detailsId}>
                      <SessionDetail session={s} />
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
