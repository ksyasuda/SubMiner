import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { WordDetailData } from '../types/stats';

export function useWordDetail(wordId: number | null) {
  const [data, setData] = useState<WordDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (wordId === null) return;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getWordDetail(wordId)
      .then(setData)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, [wordId]);

  return { data, loading, error };
}
