import { useState, useEffect, useCallback } from 'react';
import { getStatsClient } from './useStatsApi';
import type { MediaLibraryItem } from '../types/stats';

const MEDIA_LIBRARY_REFRESH_DELAY_MS = 1_500;
const MEDIA_LIBRARY_MAX_RETRIES = 3;

export function shouldRefreshMediaLibraryRows(rows: MediaLibraryItem[]): boolean {
  return rows.some((row) => {
    if (!row.youtubeVideoId) {
      return false;
    }
    return !row.videoTitle?.trim() || !row.channelName?.trim() || !row.channelThumbnailUrl?.trim();
  });
}

export function useMediaLibrary() {
  const [media, setMedia] = useState<MediaLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [version, setVersion] = useState(0);

  const refresh = useCallback(() => setVersion((v) => v + 1), []);

  useEffect(() => {
    let cancelled = false;
    let retryCount = 0;
    let retryTimer: ReturnType<typeof setTimeout> | null = null;

    const load = (isInitial = false) => {
      if (isInitial) {
        setLoading(true);
        setError(null);
      }
      getStatsClient()
        .getMediaLibrary()
        .then((rows) => {
          if (cancelled) return;
          setMedia(rows);
          if (shouldRefreshMediaLibraryRows(rows) && retryCount < MEDIA_LIBRARY_MAX_RETRIES) {
            retryCount += 1;
            retryTimer = setTimeout(() => {
              retryTimer = null;
              load(false);
            }, MEDIA_LIBRARY_REFRESH_DELAY_MS);
          }
        })
        .catch((err: Error) => {
          if (cancelled) return;
          setError(err.message);
        })
        .finally(() => {
          if (cancelled || !isInitial) return;
          setLoading(false);
        });
    };

    load(true);
    return () => {
      cancelled = true;
      if (retryTimer) {
        clearTimeout(retryTimer);
      }
    };
  }, [version]);

  return { media, loading, error, refresh };
}
