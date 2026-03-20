import { useEffect, useState } from 'react';
import { useMediaDetail } from '../../hooks/useMediaDetail';
import { apiClient } from '../../lib/api-client';
import { confirmSessionDelete } from '../../lib/delete-confirm';
import { getSessionDisplayWordCount } from '../../lib/session-word-count';
import { MediaHeader } from './MediaHeader';
import { MediaSessionList } from './MediaSessionList';
import type { SessionSummary } from '../../types/stats';

interface MediaDetailViewProps {
  videoId: number;
  initialExpandedSessionId?: number | null;
  onConsumeInitialExpandedSession?: () => void;
  onBack: () => void;
  backLabel?: string;
  onNavigateToAnime?: (animeId: number) => void;
}

export function MediaDetailView({
  videoId,
  initialExpandedSessionId = null,
  onConsumeInitialExpandedSession,
  onBack,
  backLabel = 'Back to Library',
  onNavigateToAnime,
}: MediaDetailViewProps) {
  const { data, loading, error } = useMediaDetail(videoId);
  const [localSessions, setLocalSessions] = useState<SessionSummary[] | null>(null);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [deletingSessionId, setDeletingSessionId] = useState<number | null>(null);

  useEffect(() => {
    setLocalSessions(data?.sessions ?? null);
  }, [data?.sessions]);

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data?.detail) return <div className="text-ctp-overlay2 p-4">Media not found</div>;

  const sessions = localSessions ?? data.sessions;
  const animeId = data.detail.animeId;
  const detail = {
    ...data.detail,
    totalSessions: sessions.length,
    totalActiveMs: sessions.reduce((sum, session) => sum + session.activeWatchedMs, 0),
    totalCards: sessions.reduce((sum, session) => sum + session.cardsMined, 0),
    totalTokensSeen: sessions.reduce(
      (sum, session) => sum + getSessionDisplayWordCount(session),
      0,
    ),
    totalLinesSeen: sessions.reduce((sum, session) => sum + session.linesSeen, 0),
    totalLookupCount: sessions.reduce((sum, session) => sum + session.lookupCount, 0),
    totalLookupHits: sessions.reduce((sum, session) => sum + session.lookupHits, 0),
    totalYomitanLookupCount: sessions.reduce((sum, session) => sum + session.yomitanLookupCount, 0),
  };

  const handleDeleteSession = async (session: SessionSummary) => {
    if (!confirmSessionDelete()) return;

    setDeleteError(null);
    setDeletingSessionId(session.sessionId);
    try {
      await apiClient.deleteSession(session.sessionId);
      setLocalSessions((prev) =>
        (prev ?? data.sessions).filter((item) => item.sessionId !== session.sessionId),
      );
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : 'Failed to delete session.');
    } finally {
      setDeletingSessionId(null);
    }
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <button
          type="button"
          onClick={onBack}
          className="text-sm text-ctp-blue hover:text-ctp-sapphire transition-colors"
        >
          &larr; {backLabel}
        </button>
        {onNavigateToAnime != null && animeId != null ? (
          <button
            type="button"
            onClick={() => onNavigateToAnime(animeId)}
            className="text-sm text-ctp-blue hover:text-ctp-sapphire transition-colors"
          >
            View Anime &rarr;
          </button>
        ) : null}
      </div>
      <MediaHeader detail={detail} />
      {deleteError ? <div className="text-sm text-ctp-red">{deleteError}</div> : null}
      <MediaSessionList
        sessions={sessions}
        onDeleteSession={handleDeleteSession}
        deletingSessionId={deletingSessionId}
        initialExpandedSessionId={initialExpandedSessionId}
        onConsumeInitialExpandedSession={onConsumeInitialExpandedSession}
      />
    </div>
  );
}
