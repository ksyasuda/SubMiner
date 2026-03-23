import { useState, useMemo } from 'react';
import { useMediaLibrary } from '../../hooks/useMediaLibrary';
import { formatDuration, formatNumber } from '../../lib/formatters';
import { groupMediaLibraryItems } from '../../lib/media-library-grouping';
import { CoverImage } from './CoverImage';
import { MediaCard } from './MediaCard';
import { MediaDetailView } from './MediaDetailView';

interface LibraryTabProps {
  onNavigateToSession: (sessionId: number) => void;
}

export function LibraryTab({ onNavigateToSession }: LibraryTabProps) {
  const { media, loading, error } = useMediaLibrary();
  const [search, setSearch] = useState('');
  const [selectedVideoId, setSelectedVideoId] = useState<number | null>(null);

  const filtered = useMemo(() => {
    if (!search.trim()) return media;
    const q = search.toLowerCase();
    return media.filter((m) => {
      const haystacks = [
        m.canonicalTitle,
        m.videoTitle,
        m.channelName,
        m.uploaderId,
        m.channelId,
      ].filter(Boolean);
      return haystacks.some((value) => value!.toLowerCase().includes(q));
    });
  }, [media, search]);
  const grouped = useMemo(() => groupMediaLibraryItems(filtered), [filtered]);

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
          {grouped.length} channel{grouped.length !== 1 ? 's' : ''} · {filtered.length} video
          {filtered.length !== 1 ? 's' : ''} · {formatDuration(totalMs)}
        </div>
      </div>

      {filtered.length === 0 ? (
        <div className="text-sm text-ctp-overlay2 p-4">No media found</div>
      ) : (
        <div className="space-y-6">
          {grouped.map((group) => (
            <section
              key={group.key}
              className="rounded-2xl border border-ctp-surface1 bg-ctp-surface0/70 overflow-hidden"
            >
              <div className="flex items-center gap-4 p-4 border-b border-ctp-surface1 bg-ctp-base/40">
                <CoverImage
                  videoId={group.items[0]!.videoId}
                  title={group.title}
                  src={group.imageUrl}
                  className="w-16 h-16 rounded-2xl shrink-0"
                />
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    {group.channelUrl ? (
                      <a
                        href={group.channelUrl}
                        target="_blank"
                        rel="noreferrer"
                        className="text-base font-semibold text-ctp-text truncate hover:text-ctp-blue transition-colors"
                      >
                        {group.title}
                      </a>
                    ) : (
                      <h3 className="text-base font-semibold text-ctp-text truncate">
                        {group.title}
                      </h3>
                    )}
                  </div>
                  {group.subtitle ? (
                    <div className="text-xs text-ctp-overlay1 truncate mt-1">{group.subtitle}</div>
                  ) : null}
                  <div className="text-xs text-ctp-overlay2 mt-2">
                    {group.items.length} video{group.items.length !== 1 ? 's' : ''} ·{' '}
                    {formatDuration(group.totalActiveMs)} · {formatNumber(group.totalCards)} cards
                  </div>
                </div>
              </div>
              <div className="p-4">
                <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-4">
                  {group.items.map((item) => (
                    <MediaCard
                      key={item.videoId}
                      item={item}
                      onClick={() => setSelectedVideoId(item.videoId)}
                    />
                  ))}
                </div>
              </div>
            </section>
          ))}
        </div>
      )}
    </div>
  );
}
