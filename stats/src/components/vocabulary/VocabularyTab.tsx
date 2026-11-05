import { useMemo, useState } from 'react';
import type { ExcludedWord } from '../../hooks/useExcludedWords';
import { useVocabulary } from '../../hooks/useVocabulary';
import { epochDayToDate, formatNumber } from '../../lib/formatters';
import { writeBooleanPreference } from '../../lib/preference-storage';
import type { DictionaryWordRow, KanjiEntry, StatsVocabularyCharts } from '../../types/stats';
import { SummaryTile } from '../layout/SummaryTile';
import { TrendChart } from '../trends/TrendChart';
import { DuplicateLineCleanup } from './DuplicateLineCleanup';
import { ExclusionManager } from './ExclusionManager';
import { KanjiBreakdown } from './KanjiBreakdown';
import { KanjiDetailPanel } from './KanjiDetailPanel';
import { VocabularyCoverage } from './VocabularyCoverage';
import { WordExplorer } from './WordExplorer';
import { HIDE_NAMES_KEY, loadHideNames } from './word-explorer-query';

interface VocabularyTabProps {
  onNavigateToAnime?: (animeId: number) => void;
  onOpenWordDetail?: (wordId: number) => void;
  excluded: ExcludedWord[];
  onToggleExclusion: (w: { headword: string; word: string; reading: string }) => void;
  onRemoveExclusion: (w: ExcludedWord) => void;
  onClearExclusions: () => void;
}

const GROWTH_COLOR = '#8aadf4';

/** Running total of words seen, one point per day a new word appeared. */
export function buildVocabularyGrowth(timeline: StatsVocabularyCharts['newWordsTimeline']) {
  let total = 0;
  return timeline.map((point) => {
    total += point.wordCount;
    return {
      label: epochDayToDate(point.epochDay).toLocaleDateString(undefined, {
        month: 'short',
        day: 'numeric',
      }),
      value: total,
    };
  });
}

export function VocabularyTab({
  onNavigateToAnime,
  onOpenWordDetail,
  excluded,
  onToggleExclusion,
  onRemoveExclusion,
  onClearExclusions,
}: VocabularyTabProps) {
  const {
    kanji,
    knownWords,
    summary,
    charts,
    coverage,
    dataRevision,
    loading,
    error,
    aggregatesError,
    refreshAggregates,
    reload,
  } = useVocabulary();
  const [selectedKanjiId, setSelectedKanjiId] = useState<number | null>(null);
  const [hideNames, setHideNames] = useState(() => loadHideNames());
  const [showExclusionManager, setShowExclusionManager] = useState(false);
  const [showDuplicateLineCleanup, setShowDuplicateLineCleanup] = useState(false);

  const growth = useMemo(
    () =>
      buildVocabularyGrowth(
        (hideNames ? charts?.newWordsTimelineWithoutNames : charts?.newWordsTimeline) ?? [],
      ),
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

  const updateHideNames = (hide: boolean) => {
    setHideNames(hide);
    writeBooleanPreference(HIDE_NAMES_KEY, hide);
  };
  const hideWord = (row: DictionaryWordRow) => {
    onToggleExclusion({ headword: row.headword, word: row.headword, reading: row.reading });
  };
  const openKanjiDetail = (entry: KanjiEntry) => setSelectedKanjiId(entry.kanjiId);

  const uniqueWords = (hideNames ? summary?.uniqueWordsWithoutNames : summary?.uniqueWords) ?? 0;
  const newThisWeek = (hideNames ? summary?.newThisWeekWithoutNames : summary?.newThisWeek) ?? 0;
  const knownWordCount =
    (hideNames ? summary?.knownWordCountWithoutNames : summary?.knownWordCount) ?? null;
  const knownPercent =
    knownWordCount !== null && uniqueWords > 0
      ? Math.round((knownWordCount / uniqueWords) * 100)
      : 0;

  return (
    <div className="space-y-4">
      <section
        aria-label="Vocabulary summary"
        className="grid grid-cols-2 gap-4 rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4 lg:grid-cols-4"
      >
        <SummaryTile
          label="Words seen"
          value={summary ? formatNumber(uniqueWords) : '…'}
          detail={hideNames ? 'Dictionary forms, names excluded' : 'Dictionary forms'}
        />
        {knownWordCount !== null ? (
          <SummaryTile
            label="Known in Anki"
            value={formatNumber(knownWordCount)}
            detail={`${knownPercent}% of words seen`}
          />
        ) : null}
        <SummaryTile
          label="New in the last 7 days"
          value={summary ? `+${formatNumber(newThisWeek)}` : '…'}
          detail={null}
        />
        <SummaryTile
          label="Kanji seen"
          value={summary ? formatNumber(summary.uniqueKanji) : '…'}
          detail={null}
        />
      </section>

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

      <div className="grid grid-cols-1 gap-4 xl:grid-cols-2">
        {coverage.length > 0 ? <VocabularyCoverage bands={coverage} /> : null}
        <div className={coverage.length > 0 ? '' : 'xl:col-span-2'}>
          <TrendChart title="Vocabulary Growth" data={growth} color={GROWTH_COLOR} type="line" />
        </div>
      </div>

      {charts && !charts.ready && (
        <p className="text-xs text-ctp-overlay1" role="status">
          Building vocabulary history in the background…
        </p>
      )}

      <WordExplorer
        hasKnownData={knownWords.size > 0}
        hideNames={hideNames}
        onHideNamesChange={updateHideNames}
        refreshKey={dataRevision}
        onOpenWord={(wordId) => onOpenWordDetail?.(wordId)}
        onHideWord={hideWord}
      />

      <KanjiBreakdown
        kanji={kanji}
        knownWords={knownWords}
        selectedKanjiId={selectedKanjiId}
        onSelectKanji={openKanjiDetail}
      />

      <div className="flex justify-end gap-4 text-[11px] text-ctp-overlay2">
        <button
          type="button"
          onClick={() => setShowExclusionManager(true)}
          className="transition-colors hover:text-ctp-text"
        >
          Exclusions{excluded.length > 0 && ` (${excluded.length})`}
        </button>
        <button
          type="button"
          onClick={() => setShowDuplicateLineCleanup(true)}
          className="transition-colors hover:text-ctp-text"
        >
          Clean up duplicates
        </button>
      </div>

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
