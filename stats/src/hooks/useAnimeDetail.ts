import { useState, useEffect, useCallback } from 'react';
import { getStatsClient } from './useStatsApi';
import type { AnimeDetailData } from '../types/stats';

export function useAnimeDetail(animeId: number | null) {
  const [data, setData] = useState<AnimeDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [reloadKey, setReloadKey] = useState(0);

  useEffect(() => {
    let cancelled = false;
    if (animeId === null) {
      setData(null);
      setLoading(false);
      setError(null);
      return () => {
        cancelled = true;
      };
    }
    setLoading(true);
    setError(null);
    getStatsClient()
      .getAnimeDetail(animeId)
      .then((next) => {
        if (cancelled) return;
        setData(next);
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
  }, [animeId, reloadKey]);

  const reload = useCallback(() => setReloadKey((k) => k + 1), []);

  return { data, loading, error, reload };
}
