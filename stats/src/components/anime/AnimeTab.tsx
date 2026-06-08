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
  const { anime, loading, error } = useAnimeLibrary();
  const [search, setSearch] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>('lastWatched');
  const [cardSize, setCardSize] = useState<LibraryCardSize>(() =>
    readLibraryCardSizePreference(
      getLibraryCardSizeStorage(typeof window === 'undefined' ? null : window),
    ),
  );
  const [selectedAnimeId, setSelectedAnimeId] = useState<number | null>(null);

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
        <div className="text-xs text-ctp-overlay2 shrink-0">
          {filtered.length} titles · {formatDuration(totalMs)}
        </div>
      </div>

      {filtered.length === 0 ? (
        <div className="text-sm text-ctp-overlay2 p-4">No titles found</div>
      ) : (
        <div className={`grid ${GRID_CLASSES[cardSize]} gap-4`}>
          {filtered.map((item) => (
            <AnimeCard
              key={item.animeId}
              anime={item}
              onClick={() => setSelectedAnimeId(item.animeId)}
            />
          ))}
        </div>
      )}
    </div>
  );
}
