import { Fragment, useState } from 'react';
import { formatDuration, formatNumber, formatRelativeDate } from '../../lib/formatters';
import { apiClient } from '../../lib/api-client';
import { confirmEpisodeDelete } from '../../lib/delete-confirm';
import { buildLookupRateDisplay } from '../../lib/yomitan-lookup';
import { EpisodeDetail } from './EpisodeDetail';
import { LibraryEntryPicker } from './LibraryEntryPicker';
import type { AnimeEpisode } from '../../types/stats';

/**
 * Row actions that only appear on hover. Keyboard focus and pointers with no
 * hover (touch) reveal them too, otherwise those users cannot reach the button
 * at all.
 */
const HOVER_REVEALED =
  'opacity-0 group-hover:opacity-100 focus-visible:opacity-100 [@media(hover:none)]:opacity-100';

interface EpisodeListProps {
  episodes: AnimeEpisode[];
  /** Entry these episodes currently belong to; excluded from the move picker. */
  animeId?: number;
  onEpisodeDeleted?: () => void;
  /** Fires after an episode is reassigned, so the caller can refetch. */
  onEpisodeMoved?: (removedPreviousAnime: boolean) => void;
  onOpenDetail?: (videoId: number) => void;
}

export function EpisodeList({
  episodes: initialEpisodes,
  animeId,
  onEpisodeDeleted,
  onEpisodeMoved,
  onOpenDetail,
}: EpisodeListProps) {
  const [expandedVideoId, setExpandedVideoId] = useState<number | null>(null);
  const [episodes, setEpisodes] = useState(initialEpisodes);
  const [movingEpisode, setMovingEpisode] = useState<AnimeEpisode | null>(null);
  const [moveTargetId, setMoveTargetId] = useState<number | null>(null);
  const [moveError, setMoveError] = useState<string | null>(null);

  if (episodes.length === 0) return null;

  const sorted = [...episodes].sort((a, b) => {
    if (a.episode != null && b.episode != null) return a.episode - b.episode;
    if (a.episode != null) return -1;
    if (b.episode != null) return 1;
    return 0;
  });

  const toggleWatched = async (videoId: number, currentWatched: number) => {
    const newWatched = currentWatched ? 0 : 1;
    setEpisodes((prev) =>
      prev.map((ep) => (ep.videoId === videoId ? { ...ep, watched: newWatched } : ep)),
    );
    try {
      await apiClient.setVideoWatched(videoId, newWatched === 1);
    } catch {
      setEpisodes((prev) =>
        prev.map((ep) => (ep.videoId === videoId ? { ...ep, watched: currentWatched } : ep)),
      );
    }
  };

  const handleDeleteEpisode = async (videoId: number, title: string) => {
    if (!(await confirmEpisodeDelete(title))) return;
    await apiClient.deleteVideo(videoId);
    setEpisodes((prev) => prev.filter((ep) => ep.videoId !== videoId));
    if (expandedVideoId === videoId) setExpandedVideoId(null);
    onEpisodeDeleted?.();
  };

  const handleMoveEpisode = async (videoId: number, targetAnimeId: number) => {
    setMoveTargetId(targetAnimeId);
    setMoveError(null);
    try {
      const result = await apiClient.moveVideoToAnime(videoId, targetAnimeId);
      setEpisodes((prev) => prev.filter((ep) => ep.videoId !== videoId));
      if (expandedVideoId === videoId) setExpandedVideoId(null);
      setMovingEpisode(null);
      onEpisodeMoved?.(result.removedPreviousAnime);
    } catch (err) {
      setMoveError(err instanceof Error ? err.message : 'Failed to move this episode.');
    } finally {
      setMoveTargetId(null);
    }
  };

  const watchedCount = episodes.filter((ep) => ep.watched).length;

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-ctp-text">Episodes</h3>
        <span className="text-xs text-ctp-overlay2">
          {watchedCount}/{episodes.length} watched
        </span>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-xs text-ctp-overlay2 border-b border-ctp-surface1">
              <th className="w-6 py-2 pr-1 font-medium" />
              <th className="text-left py-2 pr-3 font-medium">#</th>
              <th className="text-left py-2 pr-3 font-medium">Title</th>
              <th className="text-right py-2 pr-3 font-medium">Progress</th>
              <th className="text-right py-2 pr-3 font-medium">Watch Time</th>
              <th className="text-right py-2 pr-3 font-medium">Cards</th>
              <th className="text-right py-2 pr-3 font-medium">Lookup Rate</th>
              <th className="text-right py-2 pr-3 font-medium">Last Watched</th>
              <th className="w-28 py-2 font-medium" />
            </tr>
          </thead>
          <tbody>
            {sorted.map((ep, idx) => {
              const lookupRate = buildLookupRateDisplay(
                ep.totalYomitanLookupCount,
                ep.totalTokensSeen,
              );
              const progressPct =
                ep.durationMs > 0 && ep.endedMediaMs != null
                  ? Math.min(100, Math.round((ep.endedMediaMs / ep.durationMs) * 100))
                  : null;

              return (
                <Fragment key={ep.videoId}>
                  <tr
                    onClick={() =>
                      setExpandedVideoId(expandedVideoId === ep.videoId ? null : ep.videoId)
                    }
                    className="border-b border-ctp-surface1 last:border-0 cursor-pointer hover:bg-ctp-surface1/50 transition-colors group"
                  >
                    <td className="py-2 pr-1 text-ctp-overlay2 text-xs w-6">
                      {expandedVideoId === ep.videoId ? '\u25BC' : '\u25B6'}
                    </td>
                    <td className="py-2 pr-3 text-ctp-subtext0">{ep.episode ?? idx + 1}</td>
                    <td className="py-2 pr-3 text-ctp-text truncate max-w-[200px]">
                      {ep.canonicalTitle}
                    </td>
                    <td className="py-2 pr-3 text-right">
                      {progressPct != null ? (
                        <span
                          className={
                            progressPct >= 85
                              ? 'text-ctp-green'
                              : progressPct >= 50
                                ? 'text-ctp-peach'
                                : 'text-ctp-overlay2'
                          }
                        >
                          {progressPct}%
                        </span>
                      ) : (
                        <span className="text-ctp-overlay2">{'\u2014'}</span>
                      )}
                    </td>
                    <td className="py-2 pr-3 text-right text-ctp-blue">
                      {formatDuration(ep.totalActiveMs)}
                    </td>
                    <td className="py-2 pr-3 text-right text-ctp-cards-mined">
                      {formatNumber(ep.totalCards)}
                    </td>
                    <td className="py-2 pr-3 text-right">
                      <div className="text-ctp-sapphire">{lookupRate?.shortValue ?? '\u2014'}</div>
                      <div className="text-[11px] text-ctp-overlay2">
                        {lookupRate?.longValue ?? 'lookup rate'}
                      </div>
                    </td>
                    <td className="py-2 pr-3 text-right text-ctp-overlay2">
                      {ep.lastWatchedMs > 0 ? formatRelativeDate(ep.lastWatchedMs) : '\u2014'}
                    </td>
                    <td className="py-2 text-center w-28">
                      <div className="flex items-center justify-center gap-1">
                        {onOpenDetail ? (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              onOpenDetail(ep.videoId);
                            }}
                            className="px-2 py-1 rounded border border-ctp-surface2 text-[11px] text-ctp-blue hover:border-ctp-blue/50 hover:bg-ctp-blue/10 transition-colors"
                            title="Open episode details"
                          >
                            Details
                          </button>
                        ) : null}
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            void toggleWatched(ep.videoId, ep.watched);
                          }}
                          className={`w-5 h-5 rounded border transition-colors ${
                            ep.watched
                              ? 'bg-ctp-green border-ctp-green text-ctp-base'
                              : 'border-ctp-surface2 hover:border-ctp-overlay0 text-transparent hover:text-ctp-overlay0'
                          }`}
                          title={ep.watched ? 'Mark as unwatched' : 'Mark as watched'}
                        >
                          {'\u2713'}
                        </button>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setMoveError(null);
                            setMovingEpisode(ep);
                          }}
                          className={`w-5 h-5 rounded border border-ctp-surface2 text-transparent hover:border-ctp-blue/50 hover:text-ctp-blue focus-visible:text-ctp-blue hover:bg-ctp-blue/10 transition-colors text-xs flex items-center justify-center ${HOVER_REVEALED}`}
                          title="Move to another library entry"
                          aria-label="Move to another library entry"
                        >
                          {'\u2192'}
                        </button>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            void handleDeleteEpisode(ep.videoId, ep.canonicalTitle);
                          }}
                          className={`w-5 h-5 rounded border border-ctp-surface2 text-transparent hover:border-ctp-red/50 hover:text-ctp-red focus-visible:text-ctp-red hover:bg-ctp-red/10 transition-colors text-xs flex items-center justify-center ${HOVER_REVEALED}`}
                          title="Delete episode"
                          aria-label="Delete episode"
                        >
                          {'\u2715'}
                        </button>
                      </div>
                    </td>
                  </tr>
                  {expandedVideoId === ep.videoId && (
                    <tr>
                      <td colSpan={9} className="py-2">
                        <EpisodeDetail videoId={ep.videoId} onSessionDeleted={onEpisodeDeleted} />
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
      {movingEpisode && (
        <LibraryEntryPicker
          heading={`Move "${movingEpisode.canonicalTitle}" To`}
          excludeAnimeIds={animeId != null ? [animeId] : []}
          busyAnimeId={moveTargetId}
          error={moveError}
          onSelect={(entry) => void handleMoveEpisode(movingEpisode.videoId, entry.animeId)}
          onClose={() => {
            setMovingEpisode(null);
            setMoveError(null);
          }}
        />
      )}
    </div>
  );
}
