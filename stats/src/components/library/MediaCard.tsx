import { CoverImage } from './CoverImage';
import { formatDuration, formatNumber } from '../../lib/formatters';
import type { MediaLibraryItem } from '../../types/stats';

interface MediaCardProps {
  item: MediaLibraryItem;
  onClick: () => void;
}

export function MediaCard({ item, onClick }: MediaCardProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg overflow-hidden hover:border-ctp-surface2 transition-colors text-left w-full"
    >
      <CoverImage
        videoId={item.videoId}
        title={item.canonicalTitle}
        className="w-full aspect-[3/4] rounded-t-lg"
      />
      <div className="p-3">
        <div className="text-sm font-medium text-ctp-text truncate">{item.canonicalTitle}</div>
        <div className="text-xs text-ctp-overlay2 mt-1">
          {formatDuration(item.totalActiveMs)} · {formatNumber(item.totalCards)} cards
        </div>
        <div className="text-xs text-ctp-overlay2">
          {item.totalSessions} session{item.totalSessions !== 1 ? 's' : ''}
        </div>
      </div>
    </button>
  );
}
