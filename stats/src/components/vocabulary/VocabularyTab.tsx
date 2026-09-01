import { useState, useMemo } from 'react';
import { useVocabulary } from '../../hooks/useVocabulary';
import { StatCard } from '../layout/StatCard';
import { WordList } from './WordList';
import { KanjiBreakdown } from './KanjiBreakdown';
import { KanjiDetailPanel } from './KanjiDetailPanel';
import { ExclusionManager } from './ExclusionManager';
import { DuplicateLineCleanup } from './DuplicateLineCleanup';
import { epochDayToDate, formatNumber } from '../../lib/formatters';
import { TrendChart } from '../trends/TrendChart';
import { FrequencyRankTable } from './FrequencyRankTable';
import { CrossAnimeWordsTable } from './CrossAnimeWordsTable';
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
  const {
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
  } = useVocabulary();
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
  const chartData = useMemo(
    () => ({
      topWords: ((hideNames ? charts?.topWordsWithoutNames : charts?.topWords) ?? []).map(
        (word) => ({
          label: word.headword,
          value: word.frequency,
        }),
      ),
      newWordsTimeline: (
        (hideNames ? charts?.newWordsTimelineWithoutNames : charts?.newWordsTimeline) ?? []
      ).map((point) => ({
        label: epochDayToDate(point.epochDay).toLocaleDateString(undefined, {
          month: 'short',
          day: 'numeric',
        }),
        value: point.wordCount,
      })),
    }),
    [charts, hideNames],
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
    const match = (hideNames ? charts?.topWordsWithoutNames : charts?.topWords)?.find(
      (word) => word.headword === headword,
    );
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
          value={summary ? formatNumber(displayedSummary.uniqueWords) : '…'}
          color="text-ctp-blue"
        />
        {displayedSummary.knownWordCount !== null ? (
          <StatCard
            label="Known Words"
            value={`${formatNumber(displayedSummary.knownWordCount)} (${displayedSummary.uniqueWords > 0 ? Math.round((displayedSummary.knownWordCount / displayedSummary.uniqueWords) * 100) : 0}%)`}
            color="text-ctp-green"
          />
        ) : knownWords.size > 0 ? (
          <StatCard label="Known Words" value="…" color="text-ctp-green" />
        ) : null}
        <StatCard
          label="Unique Kanji"
          value={summary ? formatNumber(summary.uniqueKanji) : '…'}
          color="text-ctp-teal"
        />
        <StatCard
          label="New This Week"
          value={summary ? `+${formatNumber(displayedSummary.newThisWeek)}` : '…'}
          color="text-ctp-mauve"
        />
      </div>

      {aggregatesError && (
        <p className="text-xs text-ctp-red" role="alert">
          {aggregatesError}{' '}
          <button
            type="button"
            onClick={refreshAggregates}
            className="underline hover:text-ctp-text"
          >
            Retry
          </button>
        </p>
      )}

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
          data={chartData.topWords}
          color="#8aadf4"
          type="bar"
          onBarClick={handleBarClick}
        />
        <TrendChart
          title="New Words by Day"
          data={chartData.newWordsTimeline}
          color="#c6a0f6"
          type="line"
        />
      </div>

      {charts && !charts.ready && (
        <p className="text-xs text-ctp-overlay1" role="status">
          Building vocabulary history in the background…
        </p>
      )}

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
