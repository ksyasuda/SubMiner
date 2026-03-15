import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { MediaLibraryItem } from '../types/stats';

export function useMediaLibrary() {
  const [media, setMedia] = useState<MediaLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    getStatsClient()
      .getMediaLibrary()
      .then(setMedia)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  return { media, loading, error };
}
