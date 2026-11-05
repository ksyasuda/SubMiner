import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type {
  StatsTrendGroupBy,
  StatsTrendPresetRange,
  StatsTrendRange,
  TrendsDashboardData,
} from '../types/stats';

export type TimeRange = StatsTrendRange;
export type PresetTimeRange = StatsTrendPresetRange;
export type GroupBy = StatsTrendGroupBy;

/**
 * Loads the dashboard for a selection. The previous payload stays in `data` while
 * a new one loads, so controls and charts stay on screen; `refreshing` marks that
 * state and `loading` is only true before the first payload arrives.
 */
export function useTrends(
  range: TimeRange,
  groupBy: GroupBy,
  fillEmpty = true,
  title: string | null = null,
) {
  const [data, setData] = useState<TrendsDashboardData | null>(null);
  const [refreshing, setRefreshing] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const rangeKey = typeof range === 'string' ? range : `${range.startDay}:${range.endDay}`;

  useEffect(() => {
    let cancelled = false;
    setRefreshing(true);
    setError(null);
    getStatsClient()
      .getTrendsDashboard(range, groupBy, fillEmpty, title)
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
        setRefreshing(false);
      });
    return () => {
      cancelled = true;
    };
    // `rangeKey` stands in for `range`, which is a new object for each custom selection.
  }, [rangeKey, groupBy, fillEmpty, title]);

  return { data, loading: refreshing && data === null, refreshing, error };
}
