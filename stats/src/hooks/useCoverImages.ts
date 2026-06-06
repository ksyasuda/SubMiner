import { useEffect, useMemo, useState } from 'react';
import {
  buildCoverImageRequestKey,
  collectSessionCoverRequests,
  getCoverImageKey,
  mergeCoverImageData,
  type CoverImageMap,
} from '../lib/cover-images';
import { getCoverRetryDelayMs } from '../lib/cover-retry';
import type { SessionSummary } from '../types/stats';
import { getStatsClient } from './useStatsApi';

interface UseCoverImagesOptions {
  enabled?: boolean;
}

export function useCoverImages(
  sessions: SessionSummary[],
  options: UseCoverImagesOptions = {},
): CoverImageMap {
  const enabled = options.enabled ?? true;
  const requests = useMemo(() => collectSessionCoverRequests(sessions), [sessions]);
  const requestKey = useMemo(
    () => buildCoverImageRequestKey(requests.animeIds, requests.videoIds, enabled ? 1 : 0),
    [requests, enabled],
  );
  const [images, setImages] = useState<CoverImageMap>({});

  useEffect(() => {
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout> | null = null;
    let cachedImages: CoverImageMap = {};
    const client = getStatsClient();

    async function load(animeIds: number[], videoIds: number[], attempt: number): Promise<void> {
      if (animeIds.length === 0 && videoIds.length === 0) {
        return;
      }

      try {
        const data = await client.getCoverImages({ animeIds, videoIds });
        if (cancelled) return;
        cachedImages = mergeCoverImageData(cachedImages, data);
        setImages(cachedImages);
      } catch {
        if (cancelled) return;
      }

      const missingAnimeIds = animeIds.filter((id) => !cachedImages[getCoverImageKey('anime', id)]);
      const missingVideoIds = videoIds.filter((id) => !cachedImages[getCoverImageKey('media', id)]);
      if (missingAnimeIds.length === 0 && missingVideoIds.length === 0) {
        return;
      }

      timer = setTimeout(() => {
        void load(missingAnimeIds, missingVideoIds, attempt + 1);
      }, getCoverRetryDelayMs(attempt));
    }

    if (!enabled) {
      return () => {
        cancelled = true;
      };
    }

    if (requests.animeIds.length === 0 && requests.videoIds.length === 0) {
      setImages({});
      return () => {
        cancelled = true;
      };
    }

    void load(requests.animeIds, requests.videoIds, 0);

    return () => {
      cancelled = true;
      if (timer) clearTimeout(timer);
    };
  }, [requestKey]);

  return images;
}
