import type { AnimeLibraryGroup, AnimeLibraryItem } from '../types/stats';

/** What a library card shows; a single entry or a whole franchise. */
export type AnimeCardData = Pick<
  AnimeLibraryItem,
  | 'animeId'
  | 'canonicalTitle'
  | 'anilistId'
  | 'tmdbId'
  | 'mediaKind'
  | 'episodeCount'
  | 'totalActiveMs'
  | 'totalCards'
> & {
  /** Set for franchise cards with more than one part. */
  partCount?: number;
};

/** A one-part card is exactly that entry; a franchise card shows the first part's cover. */
export function groupCardData(group: AnimeLibraryGroup): AnimeCardData {
  const cover = group.parts.find((part) => part.animeId === group.coverAnimeId) ?? group.parts[0]!;
  if (group.parts.length === 1) return cover;
  return {
    animeId: cover.animeId,
    canonicalTitle: group.title,
    anilistId: cover.anilistId,
    tmdbId: cover.tmdbId,
    mediaKind: group.mediaKind,
    episodeCount: group.episodeCount,
    totalActiveMs: group.totalActiveMs,
    totalCards: group.totalCards,
    partCount: group.parts.length,
  };
}

/** Matches the franchise name or any part's title, so "Undead King" finds Overlord. */
export function groupMatchesSearch(group: AnimeLibraryGroup, search: string): boolean {
  const term = search.trim().toLowerCase();
  if (!term) return true;
  return (
    group.title.toLowerCase().includes(term) ||
    group.parts.some((part) => part.canonicalTitle.toLowerCase().includes(term))
  );
}

const FORMAT_LABELS: Record<string, string> = {
  TV: 'TV',
  TV_SHORT: 'TV Short',
  MOVIE: 'Movie',
  SPECIAL: 'Special',
  OVA: 'OVA',
  ONA: 'ONA',
  MUSIC: 'Music',
};

/** "Movie · 2017" from AniList's format and yyyymmdd start date; null when neither is known. */
export function describeAnimePart(format: string | null, startDate: number | null): string | null {
  const label = format ? (FORMAT_LABELS[format] ?? format) : null;
  const year = startDate ? String(Math.floor(startDate / 10_000)) : null;
  const parts = [label, year].filter((value): value is string => value !== null);
  return parts.length > 0 ? parts.join(' · ') : null;
}
