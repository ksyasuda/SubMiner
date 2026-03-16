import { CoverImage } from './CoverImage';
import { formatDuration, formatNumber, formatPercent } from '../../lib/formatters';
import type { MediaDetailData } from '../../types/stats';

interface MediaHeaderProps {
  detail: NonNullable<MediaDetailData['detail']>;
}

export function MediaHeader({ detail }: MediaHeaderProps) {
  const hitRate =
    detail.totalLookupCount > 0 ? detail.totalLookupHits / detail.totalLookupCount : null;
  const avgSessionMs =
    detail.totalSessions > 0 ? Math.round(detail.totalActiveMs / detail.totalSessions) : 0;

  return (
    <div className="flex gap-4">
      <CoverImage
        videoId={detail.videoId}
        title={detail.canonicalTitle}
        className="w-32 h-44 rounded-lg shrink-0"
      />
      <div className="flex-1 min-w-0">
        <h2 className="text-lg font-bold text-ctp-text truncate">{detail.canonicalTitle}</h2>
        <div className="grid grid-cols-2 gap-2 mt-3 text-sm">
          <div>
            <div className="text-ctp-blue font-medium">{formatDuration(detail.totalActiveMs)}</div>
            <div className="text-xs text-ctp-overlay2">total watch time</div>
          </div>
          <div>
            <div className="text-ctp-green font-medium">{formatNumber(detail.totalCards)}</div>
            <div className="text-xs text-ctp-overlay2">cards mined</div>
          </div>
          <div>
            <div className="text-ctp-mauve font-medium">{formatNumber(detail.totalWordsSeen)}</div>
            <div className="text-xs text-ctp-overlay2">words seen</div>
          </div>
          <div>
            <div className="text-ctp-peach font-medium">{formatPercent(hitRate)}</div>
            <div className="text-xs text-ctp-overlay2">lookup rate</div>
          </div>
          <div>
            <div className="text-ctp-text font-medium">{detail.totalSessions}</div>
            <div className="text-xs text-ctp-overlay2">sessions</div>
          </div>
          <div>
            <div className="text-ctp-text font-medium">{formatDuration(avgSessionMs)}</div>
            <div className="text-xs text-ctp-overlay2">avg session</div>
          </div>
        </div>
      </div>
    </div>
  );
}
