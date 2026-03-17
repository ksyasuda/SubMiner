import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { OverviewData, SessionSummary } from '../types/stats';

export function useOverview() {
  const [data, setData] = useState<OverviewData | null>(null);
  const [sessions, setSessions] = useState<SessionSummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    const client = getStatsClient();
    Promise.all([client.getOverview(), client.getSessions(50)])
      .then(([overview, allSessions]) => {
        if (cancelled) return;
        setData(overview);
        setSessions(allSessions);
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
  }, []);

  return { data, sessions, loading, error };
}
