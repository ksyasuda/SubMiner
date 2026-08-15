import { AnimeCoverImage } from './AnimeCoverImage';
import { formatDuration, formatNumber } from '../../lib/formatters';
import type { AnimeLibraryItem } from '../../types/stats';

interface AnimeCardProps {
  anime: AnimeLibraryItem;
  onClick: () => void;
  /** While selecting, clicking the card toggles it instead of opening it. */
  selectable?: boolean;
  selected?: boolean;
}

export function AnimeCard({
  anime,
  onClick,
  selectable = false,
  selected = false,
}: AnimeCardProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      aria-pressed={selectable ? selected : undefined}
      className={`group bg-ctp-surface0 border rounded-lg overflow-hidden hover:shadow-lg hover:shadow-ctp-blue/10 transition-all duration-200 hover:-translate-y-1 text-left w-full ${
        selected ? 'border-ctp-blue' : 'border-ctp-surface1 hover:border-ctp-blue/50'
      }`}
    >
      <div className="overflow-hidden relative">
        <AnimeCoverImage
          animeId={anime.animeId}
          title={anime.canonicalTitle}
          coverRetryToken={anime.anilistId ?? 0}
          className="w-full aspect-[3/4] rounded-t-lg transition-transform duration-200 group-hover:scale-105"
        />
        {selectable && (
          <span
            aria-hidden="true"
            className={`absolute top-2 left-2 w-5 h-5 rounded border flex items-center justify-center text-xs ${
              selected
                ? 'bg-ctp-blue border-ctp-blue text-ctp-base'
                : 'bg-ctp-crust/70 border-ctp-surface2 text-transparent'
            }`}
          >
            {'✓'}
          </span>
        )}
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
