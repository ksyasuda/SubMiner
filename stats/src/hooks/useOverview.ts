import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { OverviewData, SessionSummary } from '../types/stats';

export function useOverview() {
  const [data, setData] = useState<OverviewData | null>(null);
  const [sessions, setSessions] = useState<SessionSummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const client = getStatsClient();
    Promise.all([client.getOverview(), client.getSessions(50)])
      .then(([overview, allSessions]) => {
        setData(overview);
        setSessions(allSessions);
      })
      .catch((err) => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  return { data, sessions, loading, error };
}
