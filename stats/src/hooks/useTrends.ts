import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { TrendsDashboardData } from '../types/stats';

export type TimeRange = '7d' | '30d' | '90d' | 'all';
export type GroupBy = 'day' | 'month';

export function useTrends(range: TimeRange, groupBy: GroupBy) {
  const [data, setData] = useState<TrendsDashboardData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getTrendsDashboard(range, groupBy)
      .then((nextData) => {
        if (cancelled) return;
        setData(nextData);
      })
      .catch((err) => {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : String(err));
      })
      .finally(() => {
        if (cancelled) return;
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [range, groupBy]);

  return { data, loading, error };
}
