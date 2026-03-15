import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { AnimeDetailData } from '../types/stats';

export function useAnimeDetail(animeId: number | null) {
  const [data, setData] = useState<AnimeDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (animeId === null) return;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getAnimeDetail(animeId)
      .then(setData)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, [animeId]);

  return { data, loading, error };
}
