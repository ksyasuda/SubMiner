import { AnimeCoverImage } from './AnimeCoverImage';
import type { AnimeDetailData, AnilistEntry } from '../../types/stats';

interface AnimeHeaderProps {
  detail: AnimeDetailData['detail'];
  anilistEntries: AnilistEntry[];
  coverRetryToken?: number;
  onChangeAnilist?: () => void;
  onChangeTmdb?: () => void;
  onDeleteAnime?: () => void;
  isDeletingAnime?: boolean;
}

function AnilistButton({ entry }: { entry: AnilistEntry }) {
  const label =
    entry.season != null
      ? `Season ${entry.season}`
      : (entry.titleEnglish ?? entry.titleRomaji ?? 'AniList');

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

export function AnimeHeader({
  detail,
  anilistEntries,
  coverRetryToken = 0,
  onChangeAnilist,
  onChangeTmdb,
  onDeleteAnime,
  isDeletingAnime = false,
}: AnimeHeaderProps) {
  const isYoutube = detail.mediaKind === 'youtube';
  const altTitles = [detail.titleRomaji, detail.titleEnglish, detail.titleNative].filter(
    (t): t is string => t != null && t !== detail.canonicalTitle,
  );
  const uniqueAltTitles = [...new Set(altTitles)];

  const hasMultipleEntries = anilistEntries.length > 1;
  const isLiveAction = detail.mediaKind === 'live_action';
  const tmdbUrl =
    detail.tmdbId && detail.tmdbType
      ? `https://www.themoviedb.org/${detail.tmdbType}/${detail.tmdbId}`
      : null;
  const coverCacheToken = (detail.anilistId ?? detail.tmdbId ?? 0) * 1_000_000 + coverRetryToken;

  return (
    <div className="flex gap-4">
      <AnimeCoverImage
        animeId={detail.animeId}
        title={detail.canonicalTitle}
        coverRetryToken={coverCacheToken}
        className="w-32 h-44 rounded-lg shrink-0"
      />
      <div className="flex-1 min-w-0">
        <h2 className="text-lg font-bold text-ctp-text truncate">{detail.canonicalTitle}</h2>
        {uniqueAltTitles.length > 0 && (
          <div className="text-xs text-ctp-overlay2 mt-0.5 truncate">
            {uniqueAltTitles.join(' · ')}
          </div>
        )}
        <div className="text-sm text-ctp-subtext0 mt-2 flex items-center gap-2">
          <span>
            {isYoutube ? 'YouTube channel · ' : ''}
            {detail.episodeCount} {isYoutube ? 'video' : 'episode'}
            {detail.episodeCount !== 1 ? 's' : ''}
          </span>
          {isLiveAction && (
            <span className="px-1.5 py-0.5 rounded text-[10px] uppercase tracking-wide bg-ctp-peach/15 text-ctp-peach">
              {detail.tmdbType === 'movie' ? 'Movie' : 'Live action'}
            </span>
          )}
        </div>
        <div className="flex flex-wrap gap-1.5 mt-2">
          {tmdbUrl && (
            <a
              href={tmdbUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-blue hover:bg-ctp-surface2 hover:text-ctp-sapphire transition-colors"
            >
              View on TMDB <span className="text-[10px]">{'\u2197'}</span>
            </a>
          )}
          {!isYoutube &&
            !isLiveAction &&
            (anilistEntries.length > 0 ? (
              hasMultipleEntries ? (
                anilistEntries.map((entry) => <AnilistButton key={entry.anilistId} entry={entry} />)
              ) : (
                <a
                  href={`https://anilist.co/anime/${anilistEntries[0]!.anilistId}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-blue hover:bg-ctp-surface2 hover:text-ctp-sapphire transition-colors"
                >
                  View on AniList <span className="text-[10px]">{'\u2197'}</span>
                </a>
              )
            ) : detail.anilistId ? (
              <a
                href={`https://anilist.co/anime/${detail.anilistId}`}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-blue hover:bg-ctp-surface2 hover:text-ctp-sapphire transition-colors"
              >
                View on AniList <span className="text-[10px]">{'\u2197'}</span>
              </a>
            ) : null)}
          {!isYoutube && onChangeAnilist && (
            <button
              type="button"
              onClick={onChangeAnilist}
              title="Search AniList and manually select the correct anime entry"
              className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-overlay2 hover:bg-ctp-surface2 hover:text-ctp-subtext0 transition-colors"
            >
              {anilistEntries.length > 0 || detail.anilistId
                ? 'Change AniList Entry'
                : 'Link to AniList'}
            </button>
          )}
          {!isYoutube && onChangeTmdb && (
            <button
              type="button"
              onClick={onChangeTmdb}
              title="Search TMDB and link this title to a live-action drama or movie"
              className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-overlay2 hover:bg-ctp-surface2 hover:text-ctp-subtext0 transition-colors"
            >
              {tmdbUrl ? 'Change TMDB Title' : 'Link to TMDB'}
            </button>
          )}
          {onDeleteAnime && (
            <button
              type="button"
              onClick={onDeleteAnime}
              disabled={isDeletingAnime}
              title="Delete this title and every session and stat recorded for it"
              className="inline-flex items-center gap-1 px-2 py-1 text-xs rounded bg-ctp-surface1 text-ctp-red/80 hover:bg-ctp-red/15 hover:text-ctp-red transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {isDeletingAnime ? (
                <span
                  aria-hidden="true"
                  className="h-3 w-3 animate-spin rounded-full border-2 border-ctp-surface2 border-t-ctp-red"
                />
              ) : (
                <span aria-hidden="true">{'\u2715'}</span>
              )}
              {isDeletingAnime ? 'Deleting\u2026' : 'Delete Entry'}
            </button>
          )}
        </div>
        {detail.description && (
          <p className="text-xs text-ctp-subtext0 mt-3 line-clamp-3 leading-relaxed">
            {detail.description}
          </p>
        )}
      </div>
    </div>
  );
}
