import { AnimeCoverImage } from './AnimeCoverImage';
import type { AnimeDetailData, AnilistEntry } from '../../types/stats';

interface AnimeHeaderProps {
  detail: AnimeDetailData['detail'];
  anilistEntries: AnilistEntry[];
}

function AnilistButton({ entry }: { entry: AnilistEntry }) {
  const label = entry.season != null
    ? `Season ${entry.season}`
    : entry.titleEnglish ?? entry.titleRomaji ?? 'AniList';

  return (
    <a
      href={`https://anilist.co/anime/${entry.anilistId}`}
      target="_blank"
      rel="noopener noreferrer"
      className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-blue hover:bg-ctp-surface2 hover:text-ctp-sapphire transition-colors"
    >
      {label}
      <span className="text-[10px]">{'\u2197'}</span>
    </a>
  );
}

export function AnimeHeader({ detail, anilistEntries }: AnimeHeaderProps) {
  const altTitles = [detail.titleRomaji, detail.titleEnglish, detail.titleNative]
    .filter((t): t is string => t != null && t !== detail.canonicalTitle);
  const uniqueAltTitles = [...new Set(altTitles)];

  const hasMultipleEntries = anilistEntries.length > 1;

  return (
    <div className="flex gap-4">
      <AnimeCoverImage
        animeId={detail.animeId}
        title={detail.canonicalTitle}
        className="w-32 h-44 rounded-lg shrink-0"
      />
      <div className="flex-1 min-w-0">
        <h2 className="text-lg font-bold text-ctp-text truncate">{detail.canonicalTitle}</h2>
        {uniqueAltTitles.length > 0 && (
          <div className="text-xs text-ctp-overlay2 mt-0.5 truncate">
            {uniqueAltTitles.join(' · ')}
          </div>
        )}
        <div className="text-sm text-ctp-subtext0 mt-2">
          {detail.episodeCount} episode{detail.episodeCount !== 1 ? 's' : ''}
        </div>
        {anilistEntries.length > 0 ? (
          <div className="flex flex-wrap gap-1.5 mt-2">
            {hasMultipleEntries ? (
              anilistEntries.map((entry) => (
                <AnilistButton key={entry.anilistId} entry={entry} />
              ))
            ) : (
              <a
                href={`https://anilist.co/anime/${anilistEntries[0]!.anilistId}`}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-blue hover:bg-ctp-surface2 hover:text-ctp-sapphire transition-colors"
              >
                View on AniList <span className="text-[10px]">{'\u2197'}</span>
              </a>
            )}
          </div>
        ) : detail.anilistId ? (
          <a
            href={`https://anilist.co/anime/${detail.anilistId}`}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1 mt-2 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-blue hover:bg-ctp-surface2 hover:text-ctp-sapphire transition-colors"
          >
            View on AniList <span className="text-[10px]">{'\u2197'}</span>
          </a>
        ) : null}
      </div>
    </div>
  );
}
