import { useCallback, useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { AnimeLibraryItem } from '../types/stats';

export function useAnimeLibrary() {
  const [anime, setAnime] = useState<AnimeLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [reloadToken, setReloadToken] = useState(0);

  const reload = useCallback(() => {
    setReloadToken((token) => token + 1);
  }, []);

  useEffect(() => {
    let cancelled = false;
    getStatsClient()
      .getAnimeLibrary()
      .then((data) => {
        if (!cancelled) setAnime(data);
      })
      .catch((err: Error) => {
        if (!cancelled) setError(err.message);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken]);

  return { anime, loading, error, reload };
}
