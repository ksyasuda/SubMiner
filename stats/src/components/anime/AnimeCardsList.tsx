import { Fragment, useState } from 'react';
import { formatNumber, formatRelativeDate } from '../../lib/formatters';
import { CollapsibleSection } from './CollapsibleSection';
import { EpisodeDetail } from './EpisodeDetail';
import type { AnimeEpisode } from '../../types/stats';

interface AnimeCardsListProps {
  episodes: AnimeEpisode[];
  totalCards: number;
}

export function AnimeCardsList({ episodes, totalCards }: AnimeCardsListProps) {
  const [expandedVideoId, setExpandedVideoId] = useState<number | null>(null);

  if (totalCards === 0) {
    return (
      <CollapsibleSection title="Cards Mined (0)" defaultOpen={false}>
        <p className="text-sm text-ctp-overlay2">No cards mined from this anime yet.</p>
      </CollapsibleSection>
    );
  }

  const withCards = episodes.filter((ep) => ep.totalCards > 0);

  return (
    <CollapsibleSection title={`Cards Mined (${formatNumber(totalCards)})`} defaultOpen={false}>
      <table className="w-full text-sm">
        <thead>
          <tr className="text-xs text-ctp-overlay2 border-b border-ctp-surface1">
            <th className="w-6 py-2 pr-1 font-medium" />
            <th className="text-left py-2 pr-3 font-medium">Episode</th>
            <th className="text-right py-2 pr-3 font-medium">Cards</th>
            <th className="text-right py-2 font-medium">Last Watched</th>
          </tr>
        </thead>
        <tbody>
          {withCards.map((ep) => (
            <Fragment key={ep.videoId}>
              <tr
                onClick={() =>
                  setExpandedVideoId(expandedVideoId === ep.videoId ? null : ep.videoId)
                }
                className="border-b border-ctp-surface1 last:border-0 cursor-pointer hover:bg-ctp-surface1/50 transition-colors"
              >
                <td className="py-2 pr-1 text-ctp-overlay2 text-xs w-6">
                  {expandedVideoId === ep.videoId ? '▼' : '▶'}
                </td>
                <td className="py-2 pr-3 text-ctp-text truncate max-w-[300px]">
                  <span className="text-ctp-subtext0 mr-2">
                    {ep.episode != null ? `#${ep.episode}` : ''}
                  </span>
                  {ep.canonicalTitle}
                </td>
                <td className="py-2 pr-3 text-right text-ctp-cards-mined">
                  {formatNumber(ep.totalCards)}
                </td>
                <td className="py-2 text-right text-ctp-overlay2">
                  {ep.lastWatchedMs > 0 ? formatRelativeDate(ep.lastWatchedMs) : '\u2014'}
                </td>
              </tr>
              {expandedVideoId === ep.videoId && (
                <tr>
                  <td colSpan={4} className="py-2">
                    <EpisodeDetail videoId={ep.videoId} />
                  </td>
                </tr>
              )}
            </Fragment>
          ))}
        </tbody>
      </table>
    </CollapsibleSection>
  );
}
