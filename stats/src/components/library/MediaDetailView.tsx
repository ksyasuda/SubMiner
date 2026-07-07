import { useEffect, useRef, useState } from 'react';
import { useTranslation } from '../../i18n';
import { useMediaDetail } from '../../hooks/useMediaDetail';
import { apiClient } from '../../lib/api-client';
import { confirmSessionDelete, confirmEpisodeDelete } from '../../lib/delete-confirm';
import { getSessionDisplayWordCount } from '../../lib/session-word-count';
import { MediaHeader } from './MediaHeader';
import { MediaSessionList } from './MediaSessionList';
import type { MediaDetailData, SessionSummary } from '../../types/stats';

interface DeleteEpisodeHandlerOptions {
  videoId: number;
  title: string;
  apiClient: { deleteVideo: (id: number) => Promise<void> };
  confirmFn: (title: string) => boolean | Promise<boolean>;
  onBack: () => void;
  setDeleteError: (msg: string | null) => void;
  t: (key: string) => string;
  /**
   * Ref used to guard against reentrant delete calls synchronously. When set,
   * a subsequent invocation while the previous request is still pending is
   * ignored so clicks during the await window can't trigger duplicate deletes.
   */
  isDeletingRef?: { current: boolean };
  /** Optional React state setter so the UI can reflect the pending state. */
  setIsDeleting?: (value: boolean) => void;
}

export function buildDeleteEpisodeHandler(opts: DeleteEpisodeHandlerOptions): () => Promise<void> {
  return async () => {
    if (opts.isDeletingRef?.current) return;
    if (opts.isDeletingRef) opts.isDeletingRef.current = true;
    let confirmed = false;
    try {
      confirmed = await opts.confirmFn(opts.title);
    } catch (err) {
      if (opts.isDeletingRef) opts.isDeletingRef.current = false;
      opts.setDeleteError(
        err instanceof Error ? err.message : opts.t('stats.media.failedConfirmDelete'),
      );
      return;
    }
    if (!confirmed) {
      if (opts.isDeletingRef) opts.isDeletingRef.current = false;
      return;
    }
    opts.setIsDeleting?.(true);
    opts.setDeleteError(null);
    try {
      await opts.apiClient.deleteVideo(opts.videoId);
      opts.onBack();
    } catch (err) {
      opts.setDeleteError(
        err instanceof Error ? err.message : opts.t('stats.media.failedDeleteEpisode'),
      );
    } finally {
      if (opts.isDeletingRef) opts.isDeletingRef.current = false;
      opts.setIsDeleting?.(false);
    }
  };
}

export function getRelatedCollectionKind(detail: MediaDetailData['detail']): 'channel' | 'anime' {
  if (detail?.channelName?.trim()) {
    return 'channel';
  }
  return 'anime';
}

// Deprecated: use getRelatedCollectionKind for type-safe checks. Kept for backwards compatibility
// with existing test imports; the rendered label is now produced by t() based on the kind.
export function getRelatedCollectionLabel(detail: MediaDetailData['detail']): string {
  return getRelatedCollectionKind(detail) === 'channel' ? 'View Channel' : 'View Anime';
}

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
  const { t } = useTranslation();
  const { data, loading, error } = useMediaDetail(videoId);
  const [localSessions, setLocalSessions] = useState<SessionSummary[] | null>(null);
  const [deleteError, setDeleteError] = useState<string | null>(null);
  const [deletingSessionId, setDeletingSessionId] = useState<number | null>(null);
  const [isDeletingEpisode, setIsDeletingEpisode] = useState(false);
  const isDeletingEpisodeRef = useRef(false);
  const isDeletingSessionRef = useRef(false);

  useEffect(() => {
    setLocalSessions(data?.sessions ?? null);
  }, [data?.sessions]);

  if (loading) return <div className="text-ctp-overlay2 p-4">{t('stats.mediaDetail.loading')}</div>;
  if (error)
    return (
      <div className="text-ctp-red p-4">{t('stats.mediaDetail.error', { message: error })}</div>
    );
  if (!data?.detail)
    return <div className="text-ctp-overlay2 p-4">{t('stats.mediaDetail.notFound')}</div>;

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
  const relatedCollectionLabel = getRelatedCollectionKind(detail);

  const handleDeleteSession = async (session: SessionSummary) => {
    if (isDeletingSessionRef.current) return;
    isDeletingSessionRef.current = true;
    let confirmed = false;
    try {
      confirmed = await confirmSessionDelete();
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : t('stats.media.failedConfirmDelete'));
      isDeletingSessionRef.current = false;
      return;
    }
    if (!confirmed) {
      isDeletingSessionRef.current = false;
      return;
    }

    setDeleteError(null);
    setDeletingSessionId(session.sessionId);
    try {
      await apiClient.deleteSession(session.sessionId);
      setLocalSessions((prev) =>
        (prev ?? data.sessions).filter((item) => item.sessionId !== session.sessionId),
      );
    } catch (err) {
      setDeleteError(err instanceof Error ? err.message : t('stats.media.failedDeleteSession'));
    } finally {
      setDeletingSessionId(null);
      isDeletingSessionRef.current = false;
    }
  };

  const handleDeleteEpisode = buildDeleteEpisodeHandler({
    videoId,
    title: detail.canonicalTitle,
    apiClient,
    confirmFn: confirmEpisodeDelete,
    onBack,
    setDeleteError,
    isDeletingRef: isDeletingEpisodeRef,
    setIsDeleting: setIsDeletingEpisode,
    t,
  });

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
            {relatedCollectionLabel === 'channel'
              ? t('stats.media.viewChannel')
              : t('stats.media.viewAnime')}{' '}
            &rarr;
          </button>
        ) : null}
      </div>
      <MediaHeader
        detail={detail}
        onDeleteEpisode={handleDeleteEpisode}
        isDeletingEpisode={isDeletingEpisode}
      />
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
