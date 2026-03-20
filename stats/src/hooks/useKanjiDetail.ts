import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { KanjiDetailData } from '../types/stats';

export function useKanjiDetail(kanjiId: number | null) {
  const [data, setData] = useState<KanjiDetailData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    if (kanjiId === null) {
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
      .getKanjiDetail(kanjiId)
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
  }, [kanjiId]);

  return { data, loading, error };
}
