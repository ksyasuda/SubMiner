/**
 * Library entry classification shared by the tracker, the stats HTTP layer and
 * the stats SPA. Anime entries link to AniList; live-action entries link to TMDB.
 */
export const MEDIA_KINDS = ['anime', 'live_action'] as const;
export type MediaKind = (typeof MEDIA_KINDS)[number];

export const TMDB_MEDIA_TYPES = ['tv', 'movie'] as const;
export type TmdbMediaType = (typeof TMDB_MEDIA_TYPES)[number];

export function isMediaKind(value: unknown): value is MediaKind {
  return typeof value === 'string' && (MEDIA_KINDS as readonly string[]).includes(value);
}

export function isTmdbMediaType(value: unknown): value is TmdbMediaType {
  return typeof value === 'string' && (TMDB_MEDIA_TYPES as readonly string[]).includes(value);
}
