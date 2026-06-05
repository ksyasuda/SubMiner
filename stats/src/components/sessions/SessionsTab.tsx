import { useEffect, useMemo, useState } from 'react';
import { useSessions } from '../../hooks/useSessions';
import { SessionRow } from './SessionRow';
import { SessionDetail } from './SessionDetail';
import { DeleteProgressToast } from '../common/DeleteProgressToast';
import { apiClient } from '../../lib/api-client';
import { confirmBucketDelete, confirmSessionDelete } from '../../lib/delete-confirm';
import { formatDuration, formatNumber, formatSessionDayLabel } from '../../lib/formatters';
import { groupSessionsByVideo, type SessionBucket } from '../../lib/session-grouping';
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

export interface BucketDeleteDeps {
  bucket: SessionBucket;
  apiClient: { deleteSessions: (ids: number[]) => Promise<void> };
  confirm: (title: string, count: number) => boolean | Promise<boolean>;
  /** Called once confirmation passes, just before the delete request begins. */
  onStart?: (sessionIds: number[]) => void;
  onSuccess: (deletedIds: number[]) => void;
  onError: (message: string) => void;
}

/**
 * Build a handler that deletes every session in a bucket after confirmation.
 *
 * Extracted as a pure factory so the deletion flow can be unit-tested without
 * rendering the full SessionsTab or mocking React state.
 */
export function buildBucketDeleteHandler(deps: BucketDeleteDeps): () => Promise<void> {
  const { bucket, apiClient: client, confirm, onStart, onSuccess, onError } = deps;
  return async () => {
    const title = bucket.representativeSession.canonicalTitle ?? 'this episode';
    const ids = bucket.sessions.map((s) => s.sessionId);
    try {
      if (!(await confirm(title, ids.length))) return;
      onStart?.(ids);
      await client.deleteSessions(ids);
      onSuccess(ids);
    } catch (err) {
      onError(err instanceof Error ? err.message : 'Failed to delete sessions.');
    }
  };
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
  const [expandedBuckets, setExpandedBuckets] = useState<Set<string>>(() => new Set());
  const [search, setSearch] = useState('');
  const [visibleSessions, setVisibleSessions] = useState<SessionSummary[]>([]);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [deletingSessionIds, setDeletingSessionIds] = useState<Set<number>>(() => new Set());
  const [deletingBucketKey, setDeletingBucketKey] = useState<string | null>(null);

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

  const dayGroups = useMemo(() => groupSessionsByDay(filtered), [filtered]);

  const toggleBucket = (key: string) => {
    setExpandedBuckets((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const handleDeleteSession = async (session: SessionSummary) => {
    let confirmed = false;
    try {
      confirmed = await confirmSessionDelete();
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to confirm delete.');
      return;
    }
    if (!confirmed) return;

    setDeleteError(null);
    setDeletingSessionIds((prev) => {
      const next = new Set(prev);
      next.add(session.sessionId);
      return next;
    });
    try {
      await apiClient.deleteSession(session.sessionId);
      setVisibleSessions((prev) => prev.filter((item) => item.sessionId !== session.sessionId));
      setExpandedId((prev) => (prev === session.sessionId ? null : prev));
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete session.');
    } finally {
      setDeletingSessionIds((prev) => {
        const next = new Set(prev);
        next.delete(session.sessionId);
        return next;
      });
    }
  };

  const handleDeleteBucket = async (bucket: SessionBucket) => {
    setDeleteError(null);
    setDeletingBucketKey(bucket.key);
    const handler = buildBucketDeleteHandler({
      bucket,
      apiClient,
      confirm: confirmBucketDelete,
      onStart: (ids) =>
        setDeletingSessionIds((prev) => {
          const next = new Set(prev);
          for (const id of ids) next.add(id);
          return next;
        }),
      onSuccess: (ids) => {
        const deleted = new Set(ids);
        setVisibleSessions((prev) => prev.filter((s) => !deleted.has(s.sessionId)));
        setExpandedId((prev) => (prev != null && deleted.has(prev) ? null : prev));
        setExpandedBuckets((prev) => {
          if (!prev.has(bucket.key)) return prev;
          const next = new Set(prev);
          next.delete(bucket.key);
          return next;
        });
      },
      onError: (message) => setDeleteError(message),
    });
    try {
      await handler();
    } finally {
      setDeletingBucketKey(null);
      setDeletingSessionIds((prev) => {
        const next = new Set(prev);
        for (const session of bucket.sessions) next.delete(session.sessionId);
        return next;
      });
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

      {Array.from(dayGroups.entries()).map(([dayLabel, daySessions]) => {
        const buckets = groupSessionsByVideo(daySessions);
        return (
          <div key={dayLabel}>
            <div className="flex items-center gap-3 mb-2">
              <h3 className="text-xs font-semibold text-ctp-overlay2 uppercase tracking-widest shrink-0">
                {dayLabel}
              </h3>
              <div className="flex-1 h-px bg-gradient-to-r from-ctp-surface1 to-transparent" />
            </div>
            <div className="space-y-2">
              {buckets.map((bucket) => {
                if (bucket.sessions.length === 1) {
                  const s = bucket.sessions[0]!;
                  const detailsId = `session-details-${s.sessionId}`;
                  return (
                    <div key={bucket.key}>
                      <SessionRow
                        session={s}
                        isExpanded={expandedId === s.sessionId}
                        detailsId={detailsId}
                        onToggle={() =>
                          setExpandedId(expandedId === s.sessionId ? null : s.sessionId)
                        }
                        onDelete={() => void handleDeleteSession(s)}
                        deleteDisabled={deletingSessionIds.has(s.sessionId)}
                        onNavigateToMediaDetail={onNavigateToMediaDetail}
                      />
                      {expandedId === s.sessionId && (
                        <div id={detailsId}>
                          <SessionDetail session={s} />
                        </div>
                      )}
                    </div>
                  );
                }

                const bucketBodyId = `session-bucket-${bucket.key}`;
                const isExpanded = expandedBuckets.has(bucket.key);
                const title = bucket.representativeSession.canonicalTitle ?? 'Unknown Media';
                const deleteDisabled = deletingBucketKey === bucket.key;
                return (
                  <div key={bucket.key}>
                    <div className="relative group flex items-stretch gap-2">
                      <button
                        type="button"
                        onClick={() => toggleBucket(bucket.key)}
                        aria-expanded={isExpanded}
                        aria-controls={bucketBodyId}
                        className="flex-1 bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-3 flex items-center gap-3 hover:border-ctp-surface2 transition-colors text-left"
                      >
                        <div
                          aria-hidden="true"
                          className={`text-ctp-overlay2 text-xs shrink-0 transition-transform ${
                            isExpanded ? 'rotate-90' : ''
                          }`}
                        >
                          {'\u25B6'}
                        </div>
                        <div className="min-w-0 flex-1">
                          <div className="text-sm font-medium text-ctp-text truncate">{title}</div>
                          <div className="text-xs text-ctp-overlay2">
                            {bucket.sessions.length} session
                            {bucket.sessions.length === 1 ? '' : 's'} ·{' '}
                            {formatDuration(bucket.totalActiveMs)} active ·{' '}
                            {formatNumber(bucket.totalCardsMined)} cards
                          </div>
                        </div>
                      </button>
                      <button
                        type="button"
                        onClick={() => void handleDeleteBucket(bucket)}
                        disabled={deleteDisabled}
                        aria-label={`Delete all ${bucket.sessions.length} sessions of ${title}`}
                        title="Delete all sessions in this group"
                        className="shrink-0 w-8 rounded-lg border border-ctp-surface1 bg-ctp-surface0 text-ctp-overlay2 hover:border-ctp-red/50 hover:text-ctp-red hover:bg-ctp-red/10 transition-colors flex items-center justify-center disabled:opacity-40 disabled:cursor-not-allowed opacity-0 group-hover:opacity-100 focus:opacity-100"
                      >
                        {'\u2715'}
                      </button>
                    </div>
                    {isExpanded && (
                      <div id={bucketBodyId} className="mt-2 ml-6 space-y-2">
                        {bucket.sessions.map((s) => {
                          const detailsId = `session-details-${s.sessionId}`;
                          return (
                            <div key={s.sessionId}>
                              <SessionRow
                                session={s}
                                isExpanded={expandedId === s.sessionId}
                                detailsId={detailsId}
                                onToggle={() =>
                                  setExpandedId(expandedId === s.sessionId ? null : s.sessionId)
                                }
                                onDelete={() => void handleDeleteSession(s)}
                                deleteDisabled={deletingSessionIds.has(s.sessionId)}
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
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}

      {filtered.length === 0 && (
        <div className="text-ctp-overlay2 text-sm">
          {search.trim() ? 'No sessions matching your search.' : 'No sessions recorded yet.'}
        </div>
      )}

      <DeleteProgressToast count={deletingSessionIds.size} />
    </div>
  );
}
