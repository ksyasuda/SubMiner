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
  const raw = kind === 'channel' ? item.channelThumbnailUrl : item.videoThumbnailUrl;
  const normalized = raw?.trim() ?? '';
  return normalized.length > 0 ? normalized : null;
}

export function resolveMediaCoverApiUrl(videoId: number): string {
  return `${BASE_URL}/api/stats/media/${videoId}/cover`;
}

export function summarizeMediaLibraryGroups(groups: MediaLibraryGroup[]): {
  totalMs: number;
  totalVideos: number;
} {
  return groups.reduce(
    (summary, group) => ({
      totalMs: summary.totalMs + group.totalActiveMs,
      totalVideos: summary.totalVideos + group.items.length,
    }),
    { totalMs: 0, totalVideos: 0 },
  );
}

export function groupMediaLibraryItems(items: MediaLibraryItem[]): MediaLibraryGroup[] {
  const groups = new Map<string, MediaLibraryGroup>();

  for (const item of items) {
    const channelId = item.channelId?.trim() || null;
    const channelName = item.channelName?.trim() || null;
    const uploaderId = item.uploaderId?.trim() || null;
    const videoTitle = item.videoTitle?.trim() || null;
    const key = channelId || `video:${item.videoId}`;
    const title = channelName || uploaderId || videoTitle || item.canonicalTitle;
    const subtitle = channelId
      ? channelId
      : videoTitle && videoTitle !== item.canonicalTitle
        ? videoTitle
        : null;
    const existing = groups.get(key);
    if (existing) {
      existing.items.push(item);
      existing.totalActiveMs += item.totalActiveMs;
      existing.totalCards += item.totalCards;
      existing.lastWatchedMs = Math.max(existing.lastWatchedMs, item.lastWatchedMs);
      if (!existing.imageUrl) {
        existing.imageUrl =
          resolveMediaArtworkUrl(item, 'channel') ?? resolveMediaArtworkUrl(item, 'video');
      }
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
