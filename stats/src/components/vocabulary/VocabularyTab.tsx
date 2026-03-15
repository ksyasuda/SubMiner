import { useState } from 'react';
import { useVocabulary } from '../../hooks/useVocabulary';
import { StatCard } from '../layout/StatCard';
import { WordList } from './WordList';
import { KanjiBreakdown } from './KanjiBreakdown';
import { KanjiDetailPanel } from './KanjiDetailPanel';
import { formatNumber } from '../../lib/formatters';
import { TrendChart } from '../trends/TrendChart';
import { buildVocabularySummary } from '../../lib/dashboard-data';
import type { KanjiEntry, VocabularyEntry } from '../../types/stats';

interface VocabularyTabProps {
  onNavigateToAnime?: (animeId: number) => void;
  onOpenWordDetail?: (wordId: number) => void;
}

export function VocabularyTab({ onNavigateToAnime, onOpenWordDetail }: VocabularyTabProps) {
  const { words, kanji, loading, error } = useVocabulary();
  const [selectedKanjiId, setSelectedKanjiId] = useState<number | null>(null);
  const [search, setSearch] = useState('');

  if (loading) {
    return (
      <div className="text-ctp-overlay2 p-4" role="status" aria-live="polite">
        Loading...
      </div>
    );
  }
  if (error) {
    return (
      <div className="text-ctp-red p-4" role="alert" aria-live="assertive">
        Error: {error}
      </div>
    );
  }

  const summary = buildVocabularySummary(words, kanji);

  const handleSelectWord = (entry: VocabularyEntry): void => {
    onOpenWordDetail?.(entry.wordId);
  };

  const openKanjiDetail = (entry: KanjiEntry): void => {
    setSelectedKanjiId(entry.kanjiId);
  };

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 xl:grid-cols-3 gap-3">
        <StatCard label="Unique Words" value={formatNumber(summary.uniqueWords)} color="text-ctp-blue" />
        <StatCard label="Unique Kanji" value={formatNumber(summary.uniqueKanji)} color="text-ctp-green" />
        <StatCard
          label="New This Week"
          value={`+${formatNumber(summary.newThisWeek)}`}
          color="text-ctp-mauve"
        />
      </div>

      <div className="flex flex-wrap items-center gap-3">
        <input
          type="text"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search words..."
          className="rounded border border-ctp-surface2 bg-ctp-surface1 px-3 py-1 text-xs text-ctp-text placeholder:text-ctp-overlay0 focus:border-ctp-blue focus:outline-none focus:ring-1 focus:ring-ctp-blue"
        />
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
        <TrendChart
          title="Top Repeated Words"
          data={summary.topWords}
          color="#8aadf4"
          type="bar"
        />
        <TrendChart
          title="New Words by Day"
          data={summary.newWordsTimeline}
          color="#c6a0f6"
          type="line"
        />
      </div>

      <WordList
        words={words}
        selectedKey={null}
        onSelectWord={handleSelectWord}
        search={search}
      />

      <KanjiBreakdown
        kanji={kanji}
        selectedKanjiId={selectedKanjiId}
        onSelectKanji={openKanjiDetail}
      />

      <KanjiDetailPanel
        kanjiId={selectedKanjiId}
        onClose={() => setSelectedKanjiId(null)}
        onSelectWord={onOpenWordDetail}
        onNavigateToAnime={onNavigateToAnime}
      />
    </div>
  );
}
