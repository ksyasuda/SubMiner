import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { MediaLibraryItem } from '../types/stats';

export function useMediaLibrary() {
  const [media, setMedia] = useState<MediaLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getMediaLibrary()
      .then((rows) => {
        if (cancelled) return;
        setMedia(rows);
      })
      .catch((err: Error) => {
        if (cancelled) return;
        setError(err.message);
      })
      .finally(() => {
        if (cancelled) return;
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  return { media, loading, error };
}
