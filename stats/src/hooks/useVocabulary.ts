import { useState, useEffect } from 'react';
import { getStatsClient } from './useStatsApi';
import type { VocabularyEntry, KanjiEntry } from '../types/stats';

export function useVocabulary() {
  const [words, setWords] = useState<VocabularyEntry[]>([]);
  const [kanji, setKanji] = useState<KanjiEntry[]>([]);
  const [knownWords, setKnownWords] = useState<Set<string>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

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
  }, []);

  return { words, kanji, knownWords, loading, error };
}
