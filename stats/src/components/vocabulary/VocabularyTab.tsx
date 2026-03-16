import { useState, useMemo } from 'react';
import { useVocabulary } from '../../hooks/useVocabulary';
import { StatCard } from '../layout/StatCard';
import { WordList } from './WordList';
import { KanjiBreakdown } from './KanjiBreakdown';
import { KanjiDetailPanel } from './KanjiDetailPanel';
import { formatNumber } from '../../lib/formatters';
import { TrendChart } from '../trends/TrendChart';
import { FrequencyRankTable } from './FrequencyRankTable';
import { buildVocabularySummary } from '../../lib/dashboard-data';
import type { KanjiEntry, VocabularyEntry } from '../../types/stats';

interface VocabularyTabProps {
  onNavigateToAnime?: (animeId: number) => void;
  onOpenWordDetail?: (wordId: number) => void;
}

function isProperNoun(w: VocabularyEntry): boolean {
  return w.pos2 === '固有名詞';
}

export function VocabularyTab({ onNavigateToAnime, onOpenWordDetail }: VocabularyTabProps) {
  const { words, kanji, knownWords, loading, error } = useVocabulary();
  const [selectedKanjiId, setSelectedKanjiId] = useState<number | null>(null);
  const [search, setSearch] = useState('');
  const [hideNames, setHideNames] = useState(false);

  const hasNames = useMemo(() => words.some(isProperNoun), [words]);
  const filteredWords = useMemo(
    () => hideNames ? words.filter((w) => !isProperNoun(w)) : words,
    [words, hideNames],
  );

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

  const summary = buildVocabularySummary(filteredWords, kanji);

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

      <div className="flex items-center gap-3">
        <input
          type="text"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search words..."
          className="flex-1 bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2 text-sm text-ctp-text placeholder:text-ctp-overlay2 focus:outline-none focus:border-ctp-blue"
        />
        {hasNames && (
          <button
            type="button"
            onClick={() => setHideNames(!hideNames)}
            className={`shrink-0 px-3 py-2 rounded-lg text-xs transition-colors border ${
              hideNames
                ? 'bg-ctp-surface2 text-ctp-text border-ctp-blue/50'
                : 'bg-ctp-surface0 text-ctp-overlay2 border-ctp-surface1 hover:text-ctp-subtext0'
            }`}
          >
            Hide Names
          </button>
        )}
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

      <FrequencyRankTable words={filteredWords} knownWords={knownWords} onSelectWord={handleSelectWord} />

      <WordList
        words={filteredWords}
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
