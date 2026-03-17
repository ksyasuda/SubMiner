import { useState } from 'react';
import { BASE_URL } from '../../lib/api-client';
import { formatDuration, formatRelativeDate, formatNumber } from '../../lib/formatters';
import type { SessionSummary } from '../../types/stats';

interface SessionRowProps {
  session: SessionSummary;
  isExpanded: boolean;
  detailsId: string;
  onToggle: () => void;
  onDelete: () => void;
  deleteDisabled?: boolean;
}

function CoverThumbnail({
  animeId,
  videoId,
  title,
}: {
  animeId: number | null;
  videoId: number | null;
  title: string;
}) {
  const [failed, setFailed] = useState(false);
  const fallbackChar = title.charAt(0) || '?';

  if ((!animeId && !videoId) || failed) {
    return (
      <div className="w-10 h-14 rounded bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 text-sm font-bold shrink-0">
        {fallbackChar}
      </div>
    );
  }

  const src =
    animeId != null
      ? `${BASE_URL}/api/stats/anime/${animeId}/cover`
      : `${BASE_URL}/api/stats/media/${videoId}/cover`;

  return (
    <img
      src={src}
      alt=""
      loading="lazy"
      className="w-10 h-14 rounded object-cover shrink-0 bg-ctp-surface2"
      onError={() => setFailed(true)}
    />
  );
}

export function SessionRow({
  session,
  isExpanded,
  detailsId,
  onToggle,
  onDelete,
  deleteDisabled = false,
}: SessionRowProps) {
  return (
    <div className="relative group">
      <button
        type="button"
      onClick={onToggle}
      aria-expanded={isExpanded}
      aria-controls={detailsId}
      className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-3 pr-12 flex items-center gap-3 hover:border-ctp-surface2 transition-colors text-left"
    >
        <CoverThumbnail
          animeId={session.animeId}
          videoId={session.videoId}
          title={session.canonicalTitle ?? 'Unknown'}
        />
        <div className="min-w-0 flex-1">
          <div className="text-sm font-medium text-ctp-text truncate">
            {session.canonicalTitle ?? 'Unknown Media'}
          </div>
          <div className="text-xs text-ctp-overlay2">
            {formatRelativeDate(session.startedAtMs)} · {formatDuration(session.activeWatchedMs)}{' '}
            active
          </div>
        </div>
        <div className="flex gap-4 text-xs text-center shrink-0">
          <div>
            <div className="text-ctp-green font-medium font-mono tabular-nums">
              {formatNumber(session.cardsMined)}
            </div>
            <div className="text-ctp-overlay2">cards</div>
          </div>
          <div>
            <div className="text-ctp-mauve font-medium font-mono tabular-nums">
              {formatNumber(session.wordsSeen)}
            </div>
            <div className="text-ctp-overlay2">words</div>
          </div>
        </div>
        <div
          className={`text-ctp-blue text-xs transition-transform ${isExpanded ? 'rotate-90' : ''}`}
        >
          {'\u25B8'}
        </div>
      </button>
      <button
        type="button"
        onClick={onDelete}
        disabled={deleteDisabled}
        aria-label={`Delete session ${session.canonicalTitle ?? 'Unknown Media'}`}
        className="absolute right-3 top-1/2 -translate-y-1/2 w-5 h-5 rounded border border-ctp-surface2 text-transparent hover:border-ctp-red/50 hover:text-ctp-red hover:bg-ctp-red/10 transition-colors opacity-0 group-hover:opacity-100 focus:opacity-100 flex items-center justify-center disabled:opacity-40 disabled:cursor-not-allowed"
        title="Delete session"
      >
        {'\u2715'}
      </button>
    </div>
  );
}
