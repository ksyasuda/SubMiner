import { useState, useEffect } from 'react';
import { getStatsClient } from '../../hooks/useStatsApi';
import { formatNumber } from '../../lib/formatters';
import { CollapsibleSection } from './CollapsibleSection';
import type { AnimeWord } from '../../types/stats';

interface AnimeWordListProps {
  animeId: number;
  onNavigateToWord?: (wordId: number) => void;
}

export function AnimeWordList({ animeId, onNavigateToWord }: AnimeWordListProps) {
  const [words, setWords] = useState<AnimeWord[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    getStatsClient()
      .getAnimeWords(animeId, 50)
      .then((data) => { if (!cancelled) setWords(data); })
      .catch(() => { if (!cancelled) setWords([]); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [animeId]);

  if (loading) return <div className="text-ctp-overlay2 text-sm p-4">Loading words...</div>;
  if (words.length === 0) return null;

  return (
    <CollapsibleSection title={`Top Words (${words.length})`} defaultOpen={false}>
      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2">
        {words.map((w) => (
          <button
            key={w.wordId}
            type="button"
            onClick={() => onNavigateToWord?.(w.wordId)}
            className="bg-ctp-base border border-ctp-surface1 rounded-md p-2 hover:border-ctp-blue transition-colors cursor-pointer text-left"
          >
            <div className="text-sm font-medium text-ctp-text">{w.headword}</div>
            {w.reading && w.reading !== w.headword && (
              <div className="text-xs text-ctp-overlay2">{w.reading}</div>
            )}
            <div className="flex items-center gap-2 mt-1">
              {w.partOfSpeech && (
                <span className="text-[10px] px-1.5 py-0.5 bg-ctp-surface1 text-ctp-subtext0 rounded">
                  {w.partOfSpeech}
                </span>
              )}
              <span className="text-xs text-ctp-mauve ml-auto">{formatNumber(w.frequency)}</span>
            </div>
          </button>
        ))}
      </div>
    </CollapsibleSection>
  );
}
