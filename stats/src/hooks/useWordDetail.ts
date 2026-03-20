import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { WordDetailData } from '../types/stats';

export function useWordDetail(wordId: number | null) {
  const [data, setData] = useState<WordDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    if (wordId === null) {
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
      .getWordDetail(wordId)
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
  }, [wordId]);

  return { data, loading, error };
}
