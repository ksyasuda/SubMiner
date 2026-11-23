import { AnimeCoverImage } from './AnimeCoverImage';
import { formatDuration, formatNumber } from '../../lib/formatters';
import type { AnimeCardData } from '../../lib/anime-groups';

interface AnimeCardProps {
  anime: AnimeCardData;
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
  const unit = anime.mediaKind === 'youtube' ? 'video' : 'episode';
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
          coverRetryToken={anime.anilistId ?? anime.tmdbId ?? 0}
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
        {anime.partCount ? (
          <span className="absolute top-2 right-2 px-1.5 py-0.5 rounded text-[10px] font-medium bg-ctp-crust/80 text-ctp-text border border-ctp-surface2">
            {anime.partCount} parts
          </span>
        ) : null}
      </div>
      <div className="p-3">
        <div className="text-sm font-medium text-ctp-text truncate">{anime.canonicalTitle}</div>
        <div className="text-xs text-ctp-overlay2 mt-1">
          {anime.episodeCount} {unit}
          {anime.episodeCount !== 1 ? 's' : ''}
        </div>
        <div className="text-xs text-ctp-overlay2">
          {formatDuration(anime.totalActiveMs)} · {formatNumber(anime.totalCards)} cards
        </div>
      </div>
    </button>
  );
}
