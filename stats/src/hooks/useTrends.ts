import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type {
  DailyRollup,
  MonthlyRollup,
  EpisodesPerDay,
  NewAnimePerDay,
  WatchTimePerAnime,
  SessionSummary,
  AnimeLibraryItem,
} from '../types/stats';

export type TimeRange = '7d' | '30d' | '90d' | 'all';
export type GroupBy = 'day' | 'month';

export interface TrendsData {
  rollups: DailyRollup[] | MonthlyRollup[];
  episodesPerDay: EpisodesPerDay[];
  newAnimePerDay: NewAnimePerDay[];
  watchTimePerAnime: WatchTimePerAnime[];
  sessions: SessionSummary[];
  animeLibrary: AnimeLibraryItem[];
}

export function useTrends(range: TimeRange, groupBy: GroupBy) {
  const [data, setData] = useState<TrendsData>({
    rollups: [],
    episodesPerDay: [],
    newAnimePerDay: [],
    watchTimePerAnime: [],
    sessions: [],
    animeLibrary: [],
  });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    const client = getStatsClient();
    const limitMap: Record<TimeRange, number> = { '7d': 7, '30d': 30, '90d': 90, all: 365 };
    const limit = limitMap[range];
    const monthlyLimit = Math.max(1, Math.ceil(limit / 30));
    const sessionsLimitMap: Record<TimeRange, number> = {
      '7d': 200,
      '30d': 500,
      '90d': 500,
      all: 500,
    };

    const rollupFetcher =
      groupBy === 'month' ? client.getMonthlyRollups(monthlyLimit) : client.getDailyRollups(limit);

    Promise.all([
      rollupFetcher,
      client.getEpisodesPerDay(limit),
      client.getNewAnimePerDay(limit),
      client.getWatchTimePerAnime(limit),
      client.getSessions(sessionsLimitMap[range]),
      client.getAnimeLibrary(),
    ])
      .then(
        ([rollups, episodesPerDay, newAnimePerDay, watchTimePerAnime, sessions, animeLibrary]) => {
          if (cancelled) return;
          const now = new Date();
          const localMidnight = new Date(
            now.getFullYear(),
            now.getMonth(),
            now.getDate(),
          ).getTime();
          const cutoffMs =
            range === 'all' ? null : localMidnight - (limitMap[range] - 1) * 86_400_000;
          const filteredSessions =
            cutoffMs == null ? sessions : sessions.filter((s) => s.startedAtMs >= cutoffMs);
          setData({
            rollups,
            episodesPerDay,
            newAnimePerDay,
            watchTimePerAnime,
            sessions: filteredSessions,
            animeLibrary,
          });
        },
      )
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
