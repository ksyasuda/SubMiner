/**
 * Library entry classification shared by the tracker, the stats HTTP layer and
 * the stats SPA. Anime entries link to AniList, live-action entries link to
 * TMDB, and YouTube entries are channels grouping tracked videos.
 */
export const MEDIA_KINDS = ['anime', 'live_action', 'youtube'] as const;
export type MediaKind = (typeof MEDIA_KINDS)[number];

export const TMDB_MEDIA_TYPES = ['tv', 'movie'] as const;
export type TmdbMediaType = (typeof TMDB_MEDIA_TYPES)[number];

export function isMediaKind(value: unknown): value is MediaKind {
  return typeof value === 'string' && (MEDIA_KINDS as readonly string[]).includes(value);
}

export function isTmdbMediaType(value: unknown): value is TmdbMediaType {
  return typeof value === 'string' && (TMDB_MEDIA_TYPES as readonly string[]).includes(value);
}

/**
 * Anime and live-action entries share one title namespace: both come from the
 * filename/Jellyfin parser and an entry switches between them when it is
 * relinked from AniList to TMDB or back. YouTube channels are a separate
 * namespace, so a channel never combines with a series entry and a same-named
 * anime and channel stay separate.
 */
export function shareTitleNamespace(a: MediaKind, b: MediaKind): boolean {
  return (a === 'youtube') === (b === 'youtube');
}

/**
 * SQL predicate matching rows in the same title namespace as the bound kind
 * parameter (the SQL twin of `shareTitleNamespace`).
 */
export function sameTitleNamespaceSql(column = 'media_kind'): string {
  return `(${column} = 'youtube') = (? = 'youtube')`;
}
