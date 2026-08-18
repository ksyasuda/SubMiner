import { useState, useEffect, useCallback } from 'react';
import { getStatsClient } from './useStatsApi';
import { subscribeExcludedWordsServerSync } from './useExcludedWords';
import type {
  VocabularyEntry,
  KanjiEntry,
  StatsVocabularyCharts,
  StatsVocabularySummary,
} from '../types/stats';

const AGGREGATE_RETRY_BASE_MS = 1_000;
const AGGREGATE_RETRY_MAX_MS = 30_000;
const AGGREGATE_RETRY_LIMIT = 5;
const CHART_BACKFILL_POLL_MS = 1_000;
const CHART_BACKFILL_SLOW_POLL_MS = 5_000;
const CHART_BACKFILL_FAST_POLLS = 30;
const CHART_BACKFILL_POLL_LIMIT = 60;

function aggregateRetryDelayMs(attempt: number): number {
  return Math.min(AGGREGATE_RETRY_BASE_MS * 2 ** attempt, AGGREGATE_RETRY_MAX_MS);
}

export function useVocabulary() {
  const [words, setWords] = useState<VocabularyEntry[]>([]);
  const [kanji, setKanji] = useState<KanjiEntry[]>([]);
  const [knownWords, setKnownWords] = useState<Set<string>>(new Set());
  const [summary, setSummary] = useState<StatsVocabularySummary | null>(null);
  const [charts, setCharts] = useState<StatsVocabularyCharts | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aggregatesError, setAggregatesError] = useState<string | null>(null);
  // Bumped by `reload` after maintenance rewrites the vocabulary tables.
  const [reloadToken, setReloadToken] = useState(0);
  // Bumped independently when only the server-computed summary/charts are
  // stale, e.g. after the exclusion list changes on the server.
  const [aggregatesToken, setAggregatesToken] = useState(0);
  const refreshAggregates = useCallback(() => setAggregatesToken((token) => token + 1), []);
  const reload = useCallback(() => {
    setReloadToken((token) => token + 1);
    setAggregatesToken((token) => token + 1);
  }, []);

  useEffect(() => subscribeExcludedWordsServerSync(refreshAggregates), [refreshAggregates]);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    const client = getStatsClient();
    Promise.allSettled([client.getVocabulary(500), client.getKanji(200), client.getKnownWords()])
      .then(([wordsResult, kanjiResult, knownResult]) => {
        if (cancelled) return;
        const errors: string[] = [];

        if (wordsResult.status === 'fulfilled') {
          setWords(wordsResult.value);
        } else {
          errors.push(wordsResult.reason.message);
        }

        if (kanjiResult.status === 'fulfilled') {
          setKanji(kanjiResult.value);
        } else {
          errors.push(kanjiResult.reason.message);
        }

        if (knownResult.status === 'fulfilled') {
          setKnownWords(new Set(knownResult.value));
        }

        if (errors.length > 0) {
          setError(errors.join('; '));
        }
      })
      .finally(() => {
        if (cancelled) return;
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken]);

  useEffect(() => {
    let cancelled = false;
    setAggregatesError(null);
    const client = getStatsClient();
    const timers = new Set<ReturnType<typeof setTimeout>>();
    const schedule = (fn: () => void, delayMs: number): void => {
      const timer = setTimeout(() => {
        timers.delete(timer);
        fn();
      }, delayMs);
      timers.add(timer);
    };

    const loadSummary = (attempt: number): void => {
      void client
        .getVocabularySummary()
        .then((nextSummary) => {
          if (!cancelled) setSummary(nextSummary);
        })
        .catch((summaryError: unknown) => {
          console.error('Failed to load vocabulary summary', summaryError);
          if (cancelled) return;
          if (attempt + 1 < AGGREGATE_RETRY_LIMIT) {
            schedule(() => loadSummary(attempt + 1), aggregateRetryDelayMs(attempt));
          } else {
            setAggregatesError((previous) => previous ?? 'Vocabulary totals failed to load.');
          }
        });
    };

    const loadCharts = (attempt: number, readyPolls: number): void => {
      void client
        .getVocabularyCharts()
        .then((nextCharts) => {
          if (cancelled) return;
          setCharts(nextCharts);
          if (!nextCharts.ready) {
            const completedPolls = readyPolls + 1;
            if (completedPolls < CHART_BACKFILL_POLL_LIMIT) {
              schedule(
                () => loadCharts(0, completedPolls),
                completedPolls < CHART_BACKFILL_FAST_POLLS
                  ? CHART_BACKFILL_POLL_MS
                  : CHART_BACKFILL_SLOW_POLL_MS,
              );
            } else {
              setAggregatesError(
                (previous) =>
                  previous ?? 'Vocabulary charts are still building. Retry to check again.',
              );
            }
          }
        })
        .catch((chartError: unknown) => {
          console.error('Failed to load vocabulary charts', chartError);
          if (cancelled) return;
          if (attempt + 1 < AGGREGATE_RETRY_LIMIT) {
            schedule(() => loadCharts(attempt + 1, readyPolls), aggregateRetryDelayMs(attempt));
          } else {
            setAggregatesError((previous) => previous ?? 'Vocabulary charts failed to load.');
          }
        });
    };

    loadSummary(0);
    loadCharts(0, 0);
    return () => {
      cancelled = true;
      for (const timer of timers) clearTimeout(timer);
    };
  }, [aggregatesToken]);

  return {
    words,
    kanji,
    knownWords,
    summary,
    charts,
    loading,
    error,
    aggregatesError,
    refreshAggregates,
    reload,
  };
}
