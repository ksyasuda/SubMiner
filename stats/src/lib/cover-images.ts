import type { SessionSummary, StatsCoverImagesData } from '../types/stats';

export type CoverImageKind = 'anime' | 'media';
export type CoverImageMap = Record<string, string>;

export interface CoverImageRequest {
  animeIds: number[];
  videoIds: number[];
}

function normalizePositiveIds(ids: Iterable<number | null | undefined>): number[] {
  const uniqueIds = new Set<number>();
  for (const id of ids) {
    if (typeof id === 'number' && Number.isFinite(id) && id > 0) {
      uniqueIds.add(Math.floor(id));
    }
  }
  return Array.from(uniqueIds).sort((a, b) => a - b);
}

export function getCoverImageKey(kind: CoverImageKind, id: number): string {
  return `${kind}:${id}`;
}

export function buildCoverImageRequestKey(
  animeIds: number[],
  videoIds: number[],
  refreshToken = 0,
): string {
  return `a:${animeIds.join(',')}|m:${videoIds.join(',')}|r:${refreshToken}`;
}

export function collectSessionCoverRequests(
  sessions: Pick<SessionSummary, 'animeId' | 'videoId'>[],
): CoverImageRequest {
  const animeIds: number[] = [];
  const videoIds: number[] = [];

  for (const session of sessions) {
    if (session.animeId != null) {
      animeIds.push(session.animeId);
    } else if (session.videoId != null) {
      videoIds.push(session.videoId);
    }
  }

  return {
    animeIds: normalizePositiveIds(animeIds),
    videoIds: normalizePositiveIds(videoIds),
  };
}

export function mergeCoverImageData(
  previous: CoverImageMap,
  data: StatsCoverImagesData,
): CoverImageMap {
  const next = { ...previous };

  for (const [id, image] of Object.entries(data.anime)) {
    if (image?.dataUrl) {
      next[getCoverImageKey('anime', Number(id))] = image.dataUrl;
    }
  }

  for (const [id, image] of Object.entries(data.media)) {
    if (image?.dataUrl) {
      next[getCoverImageKey('media', Number(id))] = image.dataUrl;
    }
  }

  return next;
}

export function getCoverImageSrc(
  images: CoverImageMap,
  kind: CoverImageKind,
  id: number | null,
): string | null {
  return id == null ? null : (images[getCoverImageKey(kind, id)] ?? null);
}
