import { useState, useMemo, useEffect } from 'react';
import { useAnimeLibrary } from '../../hooks/useAnimeLibrary';
import { formatDuration } from '../../lib/formatters';
import {
  getLibraryCardSizeStorage,
  readLibraryCardSizePreference,
  type LibraryCardSize,
  writeLibraryCardSizePreference,
} from '../../lib/library-card-size';
import { AnimeCard } from './AnimeCard';
import { AnimeDetailView } from './AnimeDetailView';
import { AnimeMergeDialog } from './AnimeMergeDialog';
import { DuplicateReviewStrip } from './DuplicateReviewStrip';

type SortKey = 'lastWatched' | 'watchTime' | 'cards' | 'episodes';

const GRID_CLASSES: Record<LibraryCardSize, string> = {
  sm: 'grid-cols-5 sm:grid-cols-7 md:grid-cols-9 lg:grid-cols-11',
  md: 'grid-cols-4 sm:grid-cols-5 md:grid-cols-7 lg:grid-cols-9',
  lg: 'grid-cols-3 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-7',
};

const SORT_OPTIONS: { key: SortKey; label: string }[] = [
  { key: 'lastWatched', label: 'Last Watched' },
  { key: 'watchTime', label: 'Watch Time' },
  { key: 'cards', label: 'Cards' },
  { key: 'episodes', label: 'Episodes' },
];

function sortAnime(list: ReturnType<typeof useAnimeLibrary>['anime'], key: SortKey) {
  return [...list].sort((a, b) => {
    switch (key) {
      case 'lastWatched':
        return b.lastWatchedMs - a.lastWatchedMs;
      case 'watchTime':
        return b.totalActiveMs - a.totalActiveMs;
      case 'cards':
        return b.totalCards - a.totalCards;
      case 'episodes':
        return b.episodeCount - a.episodeCount;
    }
  });
}

interface AnimeTabProps {
  initialAnimeId?: number | null;
  onClearInitialAnime?: () => void;
  onNavigateToWord?: (wordId: number) => void;
  onOpenEpisodeDetail?: (animeId: number, videoId: number) => void;
}

export function AnimeTab({
  initialAnimeId,
  onClearInitialAnime,
  onNavigateToWord,
  onOpenEpisodeDetail,
}: AnimeTabProps) {
  const {
    anime,
    loading,
    error,
    reload,
    recommendations,
    dismissRecommendation,
    dismissingRecommendationId,
    recommendationActionError,
    clearRecommendation,
  } = useAnimeLibrary();
  const [search, setSearch] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>('lastWatched');
  const [cardSize, setCardSize] = useState<LibraryCardSize>(() =>
    readLibraryCardSizePreference(
      getLibraryCardSizeStorage(typeof window === 'undefined' ? null : window),
    ),
  );
  const [selectedAnimeId, setSelectedAnimeId] = useState<number | null>(null);
  const [selectionMode, setSelectionMode] = useState(false);
  const [checkedAnimeIds, setCheckedAnimeIds] = useState<number[]>([]);
  const [showMergeDialog, setShowMergeDialog] = useState(false);
  const [reviewRecommendationId, setReviewRecommendationId] = useState<number | null>(null);
  const [reviewAnimeIds, setReviewAnimeIds] = useState<[number, number] | null>(null);

  function toggleChecked(animeId: number): void {
    setCheckedAnimeIds((ids) =>
      ids.includes(animeId) ? ids.filter((id) => id !== animeId) : [...ids, animeId],
    );
  }

  function exitSelectionMode(): void {
    setSelectionMode(false);
    setCheckedAnimeIds([]);
    setShowMergeDialog(false);
  }

  function handleCardSizeChange(size: LibraryCardSize): void {
    setCardSize(size);
    writeLibraryCardSizePreference(
      getLibraryCardSizeStorage(typeof window === 'undefined' ? null : window),
      size,
    );
  }

  useEffect(() => {
    if (initialAnimeId != null) {
      setSelectedAnimeId(initialAnimeId);
      onClearInitialAnime?.();
    }
  }, [initialAnimeId, onClearInitialAnime]);

  const filtered = useMemo(() => {
    const base = search.trim()
      ? anime.filter((a) => a.canonicalTitle.toLowerCase().includes(search.toLowerCase()))
      : anime;
    return sortAnime(base, sortKey);
  }, [anime, search, sortKey]);

  const totalMs = anime.reduce((sum, a) => sum + a.totalActiveMs, 0);
  const checkedEntries = checkedAnimeIds
    .map((animeId) => anime.find((entry) => entry.animeId === animeId))
    .filter((entry): entry is (typeof anime)[number] => entry !== undefined);
  const hydratedRecommendations = recommendations
    .map((recommendation) => ({
      ...recommendation,
      entries: recommendation.animeIds
        .map((animeId) => anime.find((entry) => entry.animeId === animeId))
        .filter((entry): entry is (typeof anime)[number] => entry !== undefined),
    }))
    .filter((recommendation) => recommendation.entries.length >= 2);
  const activeRecommendation = hydratedRecommendations[0] ?? null;
  const reviewEntries = (reviewAnimeIds ?? [])
    .map((animeId) => anime.find((entry) => entry.animeId === animeId))
    .filter((entry): entry is (typeof anime)[number] => entry !== undefined);
  const mergeEntries = reviewRecommendationId !== null ? reviewEntries : checkedEntries;

  if (selectedAnimeId !== null) {
    return (
      <AnimeDetailView
        animeId={selectedAnimeId}
        onBack={() => setSelectedAnimeId(null)}
        onNavigateToWord={onNavigateToWord}
        onOpenEpisodeDetail={
          onOpenEpisodeDetail
            ? (videoId) => onOpenEpisodeDetail(selectedAnimeId, videoId)
            : undefined
        }
        onAnimeDeleted={reload}
        onAnilistRelinked={reload}
        onEpisodeMoved={reload}
      />
    );
  }

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;

  return (
    <div className="space-y-4">
      <div className="flex items-center gap-3">
        <input
          type="text"
          placeholder="Search library..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="flex-1 bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2 text-sm text-ctp-text placeholder:text-ctp-overlay2 focus:outline-none focus:border-ctp-blue"
        />
        <select
          value={sortKey}
          onChange={(e) => setSortKey(e.target.value as SortKey)}
          className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-2 py-2 text-sm text-ctp-text focus:outline-none focus:border-ctp-blue"
        >
          {SORT_OPTIONS.map((opt) => (
            <option key={opt.key} value={opt.key}>
              {opt.label}
            </option>
          ))}
        </select>
        <div className="flex bg-ctp-surface0 rounded-lg p-0.5 border border-ctp-surface1 shrink-0">
          {(['sm', 'md', 'lg'] as const).map((size) => (
            <button
              key={size}
              onClick={() => handleCardSizeChange(size)}
              className={`px-2 py-1 rounded-md text-xs transition-colors ${
                cardSize === size
                  ? 'bg-ctp-surface2 text-ctp-text shadow-sm'
                  : 'text-ctp-overlay2 hover:text-ctp-subtext0'
              }`}
            >
              {size === 'sm' ? '▪' : size === 'md' ? '◼' : '⬛'}
            </button>
          ))}
        </div>
        <button
          type="button"
          onClick={() => (selectionMode ? exitSelectionMode() : setSelectionMode(true))}
          title="Select several entries to merge them into one"
          className={`px-2 py-2 rounded-lg border text-xs shrink-0 transition-colors ${
            selectionMode
              ? 'bg-ctp-blue/15 border-ctp-blue/40 text-ctp-blue'
              : 'bg-ctp-surface0 border-ctp-surface1 text-ctp-overlay2 hover:text-ctp-subtext0'
          }`}
        >
          {selectionMode ? 'Cancel' : 'Select'}
        </button>
        <div className="text-xs text-ctp-overlay2 shrink-0">
          {filtered.length} titles · {formatDuration(totalMs)}
        </div>
      </div>

      {activeRecommendation ? (
        <DuplicateReviewStrip
          entries={activeRecommendation.entries}
          current={1}
          total={hydratedRecommendations.length}
          dismissing={dismissingRecommendationId === activeRecommendation.recommendationId}
          error={recommendationActionError}
          onReview={() => {
            setReviewRecommendationId(activeRecommendation.recommendationId);
            setReviewAnimeIds(activeRecommendation.animeIds);
            setShowMergeDialog(true);
          }}
          onDismiss={() => void dismissRecommendation(activeRecommendation.recommendationId)}
        />
      ) : null}

      {selectionMode && (
        <div className="flex items-center justify-between gap-3 bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2">
          <div className="text-xs text-ctp-overlay2">
            {checkedEntries.length === 0
              ? 'Pick the duplicate entries to combine'
              : `${checkedEntries.length} selected`}
          </div>
          <button
            type="button"
            disabled={checkedEntries.length < 2}
            onClick={() => setShowMergeDialog(true)}
            className="px-3 py-1.5 rounded-lg bg-ctp-blue/15 border border-ctp-blue/40 text-xs text-ctp-blue hover:bg-ctp-blue/25 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
          >
            Merge Selected
          </button>
        </div>
      )}

      {filtered.length === 0 ? (
        <div className="text-sm text-ctp-overlay2 p-4">No titles found</div>
      ) : (
        <div className={`grid ${GRID_CLASSES[cardSize]} gap-4`}>
          {filtered.map((item) => (
            <AnimeCard
              key={item.animeId}
              anime={item}
              selectable={selectionMode}
              selected={checkedAnimeIds.includes(item.animeId)}
              onClick={() =>
                selectionMode ? toggleChecked(item.animeId) : setSelectedAnimeId(item.animeId)
              }
            />
          ))}
        </div>
      )}

      {showMergeDialog && mergeEntries.length >= 2 && (
        <AnimeMergeDialog
          entries={mergeEntries}
          onClose={() => {
            setShowMergeDialog(false);
            setReviewRecommendationId(null);
            setReviewAnimeIds(null);
          }}
          onMerged={() => {
            if (reviewRecommendationId !== null) {
              clearRecommendation(reviewRecommendationId);
            }
            exitSelectionMode();
            setReviewRecommendationId(null);
            setReviewAnimeIds(null);
            reload();
          }}
        />
      )}
    </div>
  );
}
