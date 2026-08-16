import { useState, useMemo } from 'react';
import { useVocabulary } from '../../hooks/useVocabulary';
import { StatCard } from '../layout/StatCard';
import { WordList } from './WordList';
import { KanjiBreakdown } from './KanjiBreakdown';
import { KanjiDetailPanel } from './KanjiDetailPanel';
import { ExclusionManager } from './ExclusionManager';
import { DuplicateLineCleanup } from './DuplicateLineCleanup';
import { formatNumber } from '../../lib/formatters';
import { TrendChart } from '../trends/TrendChart';
import { FrequencyRankTable } from './FrequencyRankTable';
import { CrossAnimeWordsTable } from './CrossAnimeWordsTable';
import { buildVocabularySummary } from '../../lib/dashboard-data';
import type { ExcludedWord } from '../../hooks/useExcludedWords';
import type { KanjiEntry, VocabularyEntry } from '../../types/stats';

interface VocabularyTabProps {
  onNavigateToAnime?: (animeId: number) => void;
  onOpenWordDetail?: (wordId: number) => void;
  excluded: ExcludedWord[];
  isExcluded: (w: { headword: string; word: string; reading: string }) => boolean;
  onRemoveExclusion: (w: ExcludedWord) => void;
  onClearExclusions: () => void;
}

function isProperNoun(w: VocabularyEntry): boolean {
  return w.pos2 === '固有名詞';
}

export function VocabularyTab({
  onNavigateToAnime,
  onOpenWordDetail,
  excluded,
  isExcluded,
  onRemoveExclusion,
  onClearExclusions,
}: VocabularyTabProps) {
  const { words, kanji, knownWords, summary, loading, error, reload } = useVocabulary();
  const [selectedKanjiId, setSelectedKanjiId] = useState<number | null>(null);
  const [hideNames, setHideNames] = useState(false);
  const [showExclusionManager, setShowExclusionManager] = useState(false);
  const [showDuplicateLineCleanup, setShowDuplicateLineCleanup] = useState(false);

  const hasNames = useMemo(() => words.some(isProperNoun), [words]);
  const filteredWords = useMemo(() => {
    let result = words;
    if (hideNames) result = result.filter((w) => !isProperNoun(w));
    if (excluded.length > 0) result = result.filter((w) => !isExcluded(w));
    return result;
  }, [words, hideNames, excluded, isExcluded]);
  const chartSummary = useMemo(
    () => buildVocabularySummary(filteredWords, kanji),
    [filteredWords, kanji],
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

  const handleSelectWord = (entry: VocabularyEntry): void => {
    onOpenWordDetail?.(entry.wordId);
  };

  const handleBarClick = (headword: string): void => {
    const match = filteredWords.find((w) => w.headword === headword);
    if (match) onOpenWordDetail?.(match.wordId);
  };

  const openKanjiDetail = (entry: KanjiEntry): void => {
    setSelectedKanjiId(entry.kanjiId);
  };

  const displayedSummary = hideNames
    ? {
        uniqueWords: summary?.uniqueWordsWithoutNames ?? 0,
        newThisWeek: summary?.newThisWeekWithoutNames ?? 0,
        knownWordCount: summary?.knownWordCountWithoutNames ?? null,
      }
    : {
        uniqueWords: summary?.uniqueWords ?? 0,
        newThisWeek: summary?.newThisWeek ?? 0,
        knownWordCount: summary?.knownWordCount ?? null,
      };

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
        <StatCard
          label="Unique Words"
          value={formatNumber(displayedSummary.uniqueWords)}
          color="text-ctp-blue"
        />
        {displayedSummary.knownWordCount !== null && (
          <StatCard
            label="Known Words"
            value={`${formatNumber(displayedSummary.knownWordCount)} (${displayedSummary.uniqueWords > 0 ? Math.round((displayedSummary.knownWordCount / displayedSummary.uniqueWords) * 100) : 0}%)`}
            color="text-ctp-green"
          />
        )}
        <StatCard
          label="Unique Kanji"
          value={formatNumber(summary?.uniqueKanji ?? 0)}
          color="text-ctp-teal"
        />
        <StatCard
          label="New This Week"
          value={`+${formatNumber(displayedSummary.newThisWeek)}`}
          color="text-ctp-mauve"
        />
      </div>

      <div className="flex items-center justify-end gap-3">
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
        <button
          type="button"
          onClick={() => setShowDuplicateLineCleanup(true)}
          className="shrink-0 rounded-lg border border-ctp-surface1 bg-ctp-surface0 px-3 py-2 text-xs text-ctp-overlay2 transition-colors hover:text-ctp-subtext0"
        >
          Duplicates
        </button>
        <button
          type="button"
          onClick={() => setShowExclusionManager(true)}
          className={`shrink-0 px-3 py-2 rounded-lg text-xs transition-colors border ${
            excluded.length > 0
              ? 'bg-ctp-surface2 text-ctp-text border-ctp-red/50'
              : 'bg-ctp-surface0 text-ctp-overlay2 border-ctp-surface1 hover:text-ctp-subtext0'
          }`}
        >
          Exclusions{excluded.length > 0 && ` (${excluded.length})`}
        </button>
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
        <TrendChart
          title="Top Repeated Words"
          data={chartSummary.topWords}
          color="#8aadf4"
          type="bar"
          onBarClick={handleBarClick}
        />
        <TrendChart
          title="New Words by Day"
          data={chartSummary.newWordsTimeline}
          color="#c6a0f6"
          type="line"
        />
      </div>

      <FrequencyRankTable
        words={filteredWords}
        knownWords={knownWords}
        onSelectWord={handleSelectWord}
      />

      <CrossAnimeWordsTable
        words={filteredWords}
        knownWords={knownWords}
        onSelectWord={handleSelectWord}
      />

      <WordList words={filteredWords} selectedKey={null} onSelectWord={handleSelectWord} />

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

      {showExclusionManager && (
        <ExclusionManager
          excluded={excluded}
          onRemove={onRemoveExclusion}
          onClearAll={onClearExclusions}
          onClose={() => setShowExclusionManager(false)}
        />
      )}

      {showDuplicateLineCleanup && (
        <DuplicateLineCleanup
          onClose={() => setShowDuplicateLineCleanup(false)}
          onCleaned={reload}
        />
      )}
    </div>
  );
}
