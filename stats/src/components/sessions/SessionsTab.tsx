import { useState, useMemo } from 'react';
import { useSessions } from '../../hooks/useSessions';
import { SessionRow } from './SessionRow';
import { SessionDetail } from './SessionDetail';
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

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return sessions;
    return sessions.filter(
      (s) => s.canonicalTitle?.toLowerCase().includes(q),
    );
  }, [sessions, search]);

  const groups = useMemo(() => groupSessionsByDay(filtered), [filtered]);

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

      {Array.from(groups.entries()).map(([dayLabel, daySessions]) => (
        <div key={dayLabel}>
          <h3 className="text-xs font-semibold text-ctp-overlay2 uppercase tracking-wider mb-2">
            {dayLabel}
          </h3>
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
