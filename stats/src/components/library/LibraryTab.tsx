import { useState, useMemo } from 'react';
import { useMediaLibrary } from '../../hooks/useMediaLibrary';
import { formatDuration } from '../../lib/formatters';
import { MediaCard } from './MediaCard';
import { MediaDetailView } from './MediaDetailView';

export function LibraryTab() {
  const { media, loading, error } = useMediaLibrary();
  const [search, setSearch] = useState('');
  const [selectedVideoId, setSelectedVideoId] = useState<number | null>(null);

  const filtered = useMemo(() => {
    if (!search.trim()) return media;
    const q = search.toLowerCase();
    return media.filter((m) => m.canonicalTitle.toLowerCase().includes(q));
  }, [media, search]);

  const totalMs = media.reduce((sum, m) => sum + m.totalActiveMs, 0);

  if (selectedVideoId !== null) {
    return <MediaDetailView videoId={selectedVideoId} onBack={() => setSelectedVideoId(null)} />;
  }

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;

  return (
    <div className="space-y-4">
      <div className="flex items-center gap-3">
        <input
          type="text"
          placeholder="Search titles..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="flex-1 bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2 text-sm text-ctp-text placeholder:text-ctp-overlay2 focus:outline-none focus:border-ctp-blue"
        />
        <div className="text-xs text-ctp-overlay2 shrink-0">
          {filtered.length} title{filtered.length !== 1 ? 's' : ''} · {formatDuration(totalMs)}
        </div>
      </div>

      {filtered.length === 0 ? (
        <div className="text-sm text-ctp-overlay2 p-4">No media found</div>
      ) : (
        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-4">
          {filtered.map((item) => (
            <MediaCard
              key={item.videoId}
              item={item}
              onClick={() => setSelectedVideoId(item.videoId)}
            />
          ))}
        </div>
      )}
    </div>
  );
}
