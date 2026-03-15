import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { KanjiDetailData } from '../types/stats';

export function useKanjiDetail(kanjiId: number | null) {
  const [data, setData] = useState<KanjiDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (kanjiId === null) return;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getKanjiDetail(kanjiId)
      .then(setData)
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }, [kanjiId]);

  return { data, loading, error };
}
