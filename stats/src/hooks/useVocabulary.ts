import { useState, useEffect, useCallback } from 'react';
import { getStatsClient } from './useStatsApi';
import type { VocabularyEntry, KanjiEntry, StatsVocabularySummary } from '../types/stats';

export function useVocabulary() {
  const [words, setWords] = useState<VocabularyEntry[]>([]);
  const [kanji, setKanji] = useState<KanjiEntry[]>([]);
  const [knownWords, setKnownWords] = useState<Set<string>>(new Set());
  const [summary, setSummary] = useState<StatsVocabularySummary | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  // Bumped by `reload` after maintenance rewrites the vocabulary tables.
  const [reloadToken, setReloadToken] = useState(0);
  const reload = useCallback(() => setReloadToken((token) => token + 1), []);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    setSummary(null);
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
    void client
      .getVocabularySummary()
      .then((nextSummary) => {
        if (!cancelled) setSummary(nextSummary);
      })
      .catch((summaryError: unknown) => {
        console.error('Failed to load vocabulary summary', summaryError);
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken]);

  return { words, kanji, knownWords, summary, loading, error, reload };
}
