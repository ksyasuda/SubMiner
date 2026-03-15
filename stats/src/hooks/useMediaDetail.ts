import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { MediaDetailData } from '../types/stats';

export function useMediaDetail(videoId: number | null) {
  const [data, setData] = useState<MediaDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (videoId === null) return;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getMediaDetail(videoId)
      .then(setData)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, [videoId]);

  return { data, loading, error };
}
