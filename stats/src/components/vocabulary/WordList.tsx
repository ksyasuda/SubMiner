import { useMemo, useState } from 'react';
import type { VocabularyEntry } from '../../types/stats';
import { PosBadge } from './pos-helpers';

interface WordListProps {
  words: VocabularyEntry[];
  selectedKey?: string | null;
  onSelectWord?: (word: VocabularyEntry) => void;
  search?: string;
}

type SortKey = 'frequency' | 'lastSeen' | 'firstSeen';

function toWordKey(word: VocabularyEntry): string {
  return `${word.headword}\u0000${word.word}\u0000${word.reading}`;
}

const PAGE_SIZE = 100;

export function WordList({ words, selectedKey = null, onSelectWord, search = '' }: WordListProps) {
  const [sortBy, setSortBy] = useState<SortKey>('frequency');
  const [page, setPage] = useState(0);

  const titleBySort: Record<SortKey, string> = {
    frequency: 'Most Seen Words',
    lastSeen: 'Recently Seen Words',
    firstSeen: 'First Seen Words',
  };

  const filtered = useMemo(() => {
    const needle = search.trim().toLowerCase();
    if (!needle) return words;
    return words.filter(
      w => w.headword.toLowerCase().includes(needle)
        || w.word.toLowerCase().includes(needle)
        || w.reading.toLowerCase().includes(needle),
    );
  }, [words, search]);

  const sorted = useMemo(() => {
    const copy = [...filtered];
    if (sortBy === 'frequency') copy.sort((a, b) => b.frequency - a.frequency);
    else if (sortBy === 'lastSeen') copy.sort((a, b) => b.lastSeen - a.lastSeen);
    else copy.sort((a, b) => b.firstSeen - a.firstSeen);
    return copy;
  }, [filtered, sortBy]);

  const totalPages = Math.ceil(sorted.length / PAGE_SIZE);
  const paged = sorted.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);
  const maxFreq = words.reduce((max, word) => Math.max(max, word.frequency), 1);

  const getFrequencyColor = (freq: number) => {
    const ratio = freq / maxFreq;
    if (ratio > 0.5) return 'text-ctp-blue bg-ctp-blue/10';
    if (ratio > 0.2) return 'text-ctp-green bg-ctp-green/10';
    return 'text-ctp-mauve bg-ctp-mauve/10';
  };

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-ctp-text">
          {titleBySort[sortBy]}
          {search && <span className="ml-2 text-ctp-overlay1 font-normal">({filtered.length} matches)</span>}
        </h3>
        <select
          value={sortBy}
          onChange={(e) => { setSortBy(e.target.value as SortKey); setPage(0); }}
          className="text-xs bg-ctp-surface1 text-ctp-subtext0 border border-ctp-surface2 rounded px-2 py-1"
        >
          <option value="frequency">Frequency</option>
          <option value="lastSeen">Last Seen</option>
          <option value="firstSeen">First Seen</option>
        </select>
      </div>
      <div className="flex flex-wrap gap-1.5">
        {paged.map((w) => (
          <button
            type="button"
            key={toWordKey(w)}
            className={`inline-flex items-center gap-1 rounded px-2 py-0.5 text-xs transition ${
              getFrequencyColor(w.frequency)
            } ${
              selectedKey === toWordKey(w)
                ? 'ring-1 ring-ctp-blue ring-offset-1 ring-offset-ctp-surface0'
                : 'hover:ring-1 hover:ring-ctp-surface2'
            }`}
            title={`${w.word} (${w.reading}) — seen ${w.frequency}x`}
            onClick={() => onSelectWord?.(w)}
          >
            {w.headword}
            {w.partOfSpeech && (
              <PosBadge pos={w.partOfSpeech} />
            )}
            <span className="opacity-60">({w.frequency})</span>
          </button>
        ))}
      </div>
      {totalPages > 1 && (
        <div className="flex items-center justify-center gap-2 mt-3">
          <button
            type="button"
            disabled={page === 0}
            className="rounded border border-ctp-surface2 px-2 py-0.5 text-xs text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue disabled:opacity-40 disabled:cursor-not-allowed"
            onClick={() => setPage(p => p - 1)}
          >
            Prev
          </button>
          <span className="text-xs text-ctp-overlay1">
            {page + 1} / {totalPages}
          </span>
          <button
            type="button"
            disabled={page >= totalPages - 1}
            className="rounded border border-ctp-surface2 px-2 py-0.5 text-xs text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue disabled:opacity-40 disabled:cursor-not-allowed"
            onClick={() => setPage(p => p + 1)}
          >
            Next
          </button>
        </div>
      )}
    </div>
  );
}

export { toWordKey };
