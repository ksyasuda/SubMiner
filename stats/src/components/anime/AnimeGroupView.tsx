import { describeAnimePart } from '../../lib/anime-groups';
import { formatDuration, formatNumber, formatRelativeDate } from '../../lib/formatters';
import type { AnimeLibraryGroup } from '../../types/stats';
import { AnimeCoverImage } from './AnimeCoverImage';

interface AnimeGroupViewProps {
  group: AnimeLibraryGroup;
  onBack: () => void;
  onOpenPart: (animeId: number) => void;
}

/** A franchise card opened: its library entries in release order, each opening its own page. */
export function AnimeGroupView({ group, onBack, onOpenPart }: AnimeGroupViewProps) {
  const cover = group.parts.find((part) => part.animeId === group.coverAnimeId) ?? group.parts[0]!;

  return (
    <div className="space-y-4">
      <button
        type="button"
        onClick={onBack}
        className="text-sm text-ctp-blue hover:text-ctp-sapphire transition-colors"
      >
        &larr; Back to Library
      </button>
      <div className="flex gap-4">
        <AnimeCoverImage
          animeId={cover.animeId}
          title={group.title}
          coverRetryToken={cover.anilistId ?? 0}
          className="w-32 h-44 rounded-lg shrink-0"
        />
        <div className="flex-1 min-w-0">
          <h2 className="text-lg font-bold text-ctp-text truncate">{group.title}</h2>
          <div className="text-sm text-ctp-subtext0 mt-2">
            {group.parts.length} parts · {group.episodeCount} episode
            {group.episodeCount !== 1 ? 's' : ''}
          </div>
          <div className="text-xs text-ctp-overlay2 mt-1">
            {formatDuration(group.totalActiveMs)} · {formatNumber(group.totalCards)} cards
            {group.lastWatchedMs > 0
              ? ` · last watched ${formatRelativeDate(group.lastWatchedMs)}`
              : ''}
          </div>
          <p className="text-xs text-ctp-overlay2 mt-3 leading-relaxed">
            Grouped by AniList&apos;s season, movie and side-story links. Open a part to move it to
            another anime or show it on its own.
          </p>
        </div>
      </div>
      <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
        <h3 className="text-sm font-semibold text-ctp-text mb-3">Seasons and Movies</h3>
        <ul className="space-y-1">
          {group.parts.map((part) => {
            const kind = describeAnimePart(part.format, part.startDate);
            return (
              <li key={part.animeId}>
                <button
                  type="button"
                  onClick={() => onOpenPart(part.animeId)}
                  className="w-full flex items-center gap-3 p-2.5 rounded-lg hover:bg-ctp-surface1 transition-colors text-left"
                >
                  <AnimeCoverImage
                    animeId={part.animeId}
                    title={part.canonicalTitle}
                    coverRetryToken={part.anilistId ?? part.tmdbId ?? 0}
                    className="w-10 h-14 rounded shrink-0"
                  />
                  <div className="min-w-0 flex-1">
                    <div className="text-sm text-ctp-text truncate">{part.canonicalTitle}</div>
                    <div className="text-xs text-ctp-overlay2 mt-0.5">
                      {kind ? `${kind} · ` : ''}
                      {part.episodeCount} episode{part.episodeCount !== 1 ? 's' : ''}
                    </div>
                  </div>
                  <div className="text-xs text-ctp-overlay2 text-right shrink-0">
                    <div>{formatDuration(part.totalActiveMs)}</div>
                    <div>{formatNumber(part.totalCards)} cards</div>
                  </div>
                </button>
              </li>
            );
          })}
        </ul>
      </div>
    </div>
  );
}
