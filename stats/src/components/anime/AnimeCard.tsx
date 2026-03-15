import { AnimeCoverImage } from './AnimeCoverImage';
import { formatDuration, formatNumber } from '../../lib/formatters';
import type { AnimeLibraryItem } from '../../types/stats';

interface AnimeCardProps {
  anime: AnimeLibraryItem;
  onClick: () => void;
}

export function AnimeCard({ anime, onClick }: AnimeCardProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="group bg-ctp-surface0 border border-ctp-surface1 rounded-lg overflow-hidden hover:border-ctp-blue/50 hover:shadow-lg hover:shadow-ctp-blue/10 transition-all duration-200 hover:-translate-y-1 text-left w-full"
    >
      <div className="overflow-hidden">
        <AnimeCoverImage
          animeId={anime.animeId}
          title={anime.canonicalTitle}
          className="w-full aspect-[3/4] rounded-t-lg transition-transform duration-200 group-hover:scale-105"
        />
      </div>
      <div className="p-3">
        <div className="text-sm font-medium text-ctp-text truncate">{anime.canonicalTitle}</div>
        <div className="text-xs text-ctp-overlay2 mt-1">
          {anime.episodeCount} episode{anime.episodeCount !== 1 ? 's' : ''}
        </div>
        <div className="text-xs text-ctp-overlay2">
          {formatDuration(anime.totalActiveMs)} · {formatNumber(anime.totalCards)} cards
        </div>
      </div>
    </button>
  );
}
