import { BASE_URL } from './api-client';
import type { MediaLibraryItem } from '../types/stats';

export interface MediaLibraryGroup {
  key: string;
  title: string;
  subtitle: string | null;
  imageUrl: string | null;
  channelUrl: string | null;
  items: MediaLibraryItem[];
  totalActiveMs: number;
  totalCards: number;
  lastWatchedMs: number;
}

export function resolveMediaArtworkUrl(
  item: Pick<MediaLibraryItem, 'videoThumbnailUrl' | 'channelThumbnailUrl'>,
  kind: 'video' | 'channel',
): string | null {
  if (kind === 'channel') {
    return item.channelThumbnailUrl ?? null;
  }
  return item.videoThumbnailUrl ?? null;
}

export function resolveMediaCoverApiUrl(videoId: number): string {
  return `${BASE_URL}/api/stats/media/${videoId}/cover`;
}

export function groupMediaLibraryItems(items: MediaLibraryItem[]): MediaLibraryGroup[] {
  const groups = new Map<string, MediaLibraryGroup>();

  for (const item of items) {
    const key = item.channelId?.trim() || `video:${item.videoId}`;
    const title =
      item.channelName?.trim() ||
      item.uploaderId?.trim() ||
      item.videoTitle?.trim() ||
      item.canonicalTitle;
    const subtitle =
      item.channelId?.trim() != null && item.channelId?.trim() !== ''
        ? `${item.channelId}`
        : item.videoTitle?.trim() && item.videoTitle !== item.canonicalTitle
          ? item.videoTitle
          : null;
    const existing = groups.get(key);
    if (existing) {
      existing.items.push(item);
      existing.totalActiveMs += item.totalActiveMs;
      existing.totalCards += item.totalCards;
      existing.lastWatchedMs = Math.max(existing.lastWatchedMs, item.lastWatchedMs);
      continue;
    }

    groups.set(key, {
      key,
      title,
      subtitle,
      imageUrl: resolveMediaArtworkUrl(item, 'channel') ?? resolveMediaArtworkUrl(item, 'video'),
      channelUrl: item.channelUrl ?? null,
      items: [item],
      totalActiveMs: item.totalActiveMs,
      totalCards: item.totalCards,
      lastWatchedMs: item.lastWatchedMs,
    });
  }

  return [...groups.values()]
    .map((group) => ({
      ...group,
      items: [...group.items].sort((a, b) => b.lastWatchedMs - a.lastWatchedMs),
    }))
    .sort((a, b) => b.lastWatchedMs - a.lastWatchedMs);
}
