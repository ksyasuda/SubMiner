import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { StreakCalendarDay } from '../types/stats';

export function useStreakCalendar(days = 90, refreshKey = 0) {
  const [calendar, setCalendar] = useState<StreakCalendarDay[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    getStatsClient()
      .getStreakCalendar(days)
      .then((data) => {
        if (!cancelled) setCalendar(data);
      })
      .catch((err: Error) => {
        if (!cancelled) setError(err.message);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [days, refreshKey]);

  return { calendar, loading, error };
}
