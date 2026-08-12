import { useCallback, useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { AnimeLibraryItem, StatsAnimeMergeRecommendation } from '../types/stats';

const BACKGROUND_REFRESH_MS = 30_000;

export function useAnimeLibrary() {
  const [anime, setAnime] = useState<AnimeLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [recommendations, setRecommendations] = useState<StatsAnimeMergeRecommendation[]>([]);
  const [dismissingRecommendationId, setDismissingRecommendationId] = useState<number | null>(null);
  const [recommendationActionError, setRecommendationActionError] = useState<string | null>(null);
  const [reloadToken, setReloadToken] = useState(0);

  const reload = useCallback(() => {
    setReloadToken((token) => token + 1);
  }, []);

  useEffect(() => {
    let cancelled = false;
    const client = getStatsClient();
    client
      .getAnimeLibrary()
      .then((data) => {
        if (!cancelled) {
          setAnime(data);
          setError(null);
        }
      })
      .catch((err: Error) => {
        if (!cancelled) setError(err.message);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });

    // Recommendation support is deliberately non-blocking. An older backend
    // should still be able to display its library even when this endpoint is
    // unavailable.
    client
      .getAnimeMergeRecommendations()
      .then((data) => {
        if (!cancelled) {
          setRecommendations(data.recommendations);
          setRecommendationActionError(null);
        }
      })
      .catch(() => {
        // Preserve the last confirmed set. A transient polling failure should
        // not make a pending review silently disappear.
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken]);

  useEffect(() => {
    const refreshOnFocus = () => reload();
    const interval = window.setInterval(reload, BACKGROUND_REFRESH_MS);
    window.addEventListener('focus', refreshOnFocus);
    return () => {
      window.clearInterval(interval);
      window.removeEventListener('focus', refreshOnFocus);
    };
  }, [reload]);

  const dismissRecommendation = useCallback(async (recommendationId: number) => {
    setDismissingRecommendationId(recommendationId);
    setRecommendationActionError(null);
    try {
      await getStatsClient().dismissAnimeMergeRecommendation(recommendationId);
      setRecommendations((current) =>
        current.filter((item) => item.recommendationId !== recommendationId),
      );
    } catch {
      setRecommendationActionError('Could not dismiss this suggestion. Try again.');
    } finally {
      setDismissingRecommendationId(null);
    }
  }, []);

  const clearRecommendation = useCallback((recommendationId: number) => {
    setRecommendations((current) =>
      current.filter((item) => item.recommendationId !== recommendationId),
    );
    setRecommendationActionError(null);
  }, []);

  return {
    anime,
    loading,
    error,
    reload,
    recommendations,
    dismissRecommendation,
    dismissingRecommendationId,
    recommendationActionError,
    clearRecommendation,
  };
}
