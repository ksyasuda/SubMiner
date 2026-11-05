import { useEffect, useState } from 'react';
import { getStatsClient } from './useStatsApi';
import type { DictionaryWordPage, StatsDictionaryWordQuery } from '../types/stats';

/**
 * Loads one page of the word list. The previous page stays in `page` while the
 * next one loads, so the table doesn't collapse between filters.
 */
export function useDictionaryWords(query: StatsDictionaryWordQuery, refreshKey: number) {
  const [page, setPage] = useState<DictionaryWordPage | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const queryKey = JSON.stringify(query);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getDictionaryWords(query)
      .then((nextPage) => {
        if (!cancelled) setPage(nextPage);
      })
      .catch((err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
    // `queryKey` stands in for `query`, which is a new object each render.
  }, [queryKey, refreshKey]);

  return { page, loading, error };
}
