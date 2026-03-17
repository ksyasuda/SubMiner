import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { MediaDetailData } from '../types/stats';

export function useMediaDetail(videoId: number | null) {
  const [data, setData] = useState<MediaDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    if (videoId === null) {
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
      .getMediaDetail(videoId)
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
  }, [videoId]);

  return { data, loading, error };
}
