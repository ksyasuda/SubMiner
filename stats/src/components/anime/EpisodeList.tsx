import { Fragment, useState } from 'react';
import { formatDuration, formatNumber, formatRelativeDate } from '../../lib/formatters';
import { apiClient } from '../../lib/api-client';
import { EpisodeDetail } from './EpisodeDetail';
import type { AnimeEpisode } from '../../types/stats';

interface EpisodeListProps {
  episodes: AnimeEpisode[];
}

export function EpisodeList({ episodes: initialEpisodes }: EpisodeListProps) {
  const [expandedVideoId, setExpandedVideoId] = useState<number | null>(null);
  const [episodes, setEpisodes] = useState(initialEpisodes);

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
              <th className="text-right py-2 pr-3 font-medium">Last Watched</th>
              <th className="w-8 py-2 font-medium" />
            </tr>
          </thead>
          <tbody>
            {sorted.map((ep, idx) => (
              <Fragment key={ep.videoId}>
                <tr
                  onClick={() => setExpandedVideoId(expandedVideoId === ep.videoId ? null : ep.videoId)}
                  className="border-b border-ctp-surface1 last:border-0 cursor-pointer hover:bg-ctp-surface1/50 transition-colors"
                >
                  <td className="py-2 pr-1 text-ctp-overlay2 text-xs w-6">
                    {expandedVideoId === ep.videoId ? '\u25BC' : '\u25B6'}
                  </td>
                  <td className="py-2 pr-3 text-ctp-subtext0">
                    {ep.episode ?? idx + 1}
                  </td>
                  <td className="py-2 pr-3 text-ctp-text truncate max-w-[200px]">
                    {ep.canonicalTitle}
                  </td>
                  <td className="py-2 pr-3 text-right">
                    {ep.durationMs > 0 ? (
                      <span className={
                        ep.totalActiveMs >= ep.durationMs * 0.85
                          ? 'text-ctp-green'
                          : ep.totalActiveMs >= ep.durationMs * 0.5
                            ? 'text-ctp-peach'
                            : 'text-ctp-overlay2'
                      }>
                        {Math.min(100, Math.round((ep.totalActiveMs / ep.durationMs) * 100))}%
                      </span>
                    ) : (
                      <span className="text-ctp-overlay2">{'\u2014'}</span>
                    )}
                  </td>
                  <td className="py-2 pr-3 text-right text-ctp-blue">
                    {formatDuration(ep.totalActiveMs)}
                  </td>
                  <td className="py-2 pr-3 text-right text-ctp-green">
                    {formatNumber(ep.totalCards)}
                  </td>
                  <td className="py-2 pr-3 text-right text-ctp-overlay2">
                    {ep.lastWatchedMs > 0 ? formatRelativeDate(ep.lastWatchedMs) : '\u2014'}
                  </td>
                  <td className="py-2 text-center w-8">
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
                  </td>
                </tr>
                {expandedVideoId === ep.videoId && (
                  <tr>
                    <td colSpan={8} className="py-2">
                      <EpisodeDetail videoId={ep.videoId} />
                    </td>
                  </tr>
                )}
              </Fragment>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
