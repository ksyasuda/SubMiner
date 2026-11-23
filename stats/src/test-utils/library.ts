import type { apiClient } from '../lib/api-client';
import type { AnimeLibraryGroup, AnimeLibraryItem } from '../types/stats';

/** Wraps each entry in a one-part card, as the server does for entries without AniList relations. */
export function asSingleGroups(items: AnimeLibraryItem[]): AnimeLibraryGroup[] {
  return items.map((item) => ({
    groupKey: -item.animeId,
    title: item.canonicalTitle,
    mediaKind: item.mediaKind,
    coverAnimeId: item.animeId,
    totalSessions: item.totalSessions,
    totalActiveMs: item.totalActiveMs,
    totalCards: item.totalCards,
    totalTokensSeen: item.totalTokensSeen,
    episodeCount: item.episodeCount,
    lastWatchedMs: item.lastWatchedMs,
    parts: [{ ...item, format: null, startDate: null }],
  }));
}

/** A `getAnimeLibraryGroups` mock serving `load()`'s entries as one-part cards. */
export function singleGroupLibrary(
  load: () => Promise<AnimeLibraryItem[]>,
): typeof apiClient.getAnimeLibraryGroups {
  return async () => asSingleGroups(await load());
}
