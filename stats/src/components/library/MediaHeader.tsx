import { useState, useEffect } from 'react';
import { CoverImage } from './CoverImage';
import { formatDuration, formatNumber, formatPercent } from '../../lib/formatters';
import { getStatsClient } from '../../hooks/useStatsApi';
import { buildLookupRateDisplay } from '../../lib/yomitan-lookup';
import type { MediaDetailData } from '../../types/stats';

interface MediaHeaderProps {
  detail: NonNullable<MediaDetailData['detail']>;
  initialKnownWordsSummary?: {
    totalUniqueWords: number;
    knownWordCount: number;
  } | null;
}

export function MediaHeader({ detail, initialKnownWordsSummary = null }: MediaHeaderProps) {
  const knownTokenRate =
    detail.totalLookupCount > 0 ? detail.totalLookupHits / detail.totalLookupCount : null;
  const avgSessionMs =
    detail.totalSessions > 0 ? Math.round(detail.totalActiveMs / detail.totalSessions) : 0;
  const lookupRate = buildLookupRateDisplay(detail.totalYomitanLookupCount, detail.totalTokensSeen);

  const [knownWordsSummary, setKnownWordsSummary] = useState<{
    totalUniqueWords: number;
    knownWordCount: number;
  } | null>(initialKnownWordsSummary);

  useEffect(() => {
    let cancelled = false;
    getStatsClient()
      .getMediaKnownWordsSummary(detail.videoId)
      .then((data) => {
        if (!cancelled) setKnownWordsSummary(data);
      })
      .catch(() => {
        if (!cancelled) setKnownWordsSummary(null);
      });
    return () => {
      cancelled = true;
    };
  }, [detail.videoId]);

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
            <div className="text-ctp-cards-mined font-medium">{formatNumber(detail.totalCards)}</div>
            <div className="text-xs text-ctp-overlay2">cards mined</div>
          </div>
          <div>
            <div className="text-ctp-mauve font-medium">{formatNumber(detail.totalTokensSeen)}</div>
            <div className="text-xs text-ctp-overlay2">token occurrences</div>
          </div>
          <div>
            <div className="text-ctp-lavender font-medium">
              {formatNumber(detail.totalYomitanLookupCount)}
            </div>
            <div className="text-xs text-ctp-overlay2">Yomitan lookups</div>
          </div>
          <div>
            <div className="text-ctp-sapphire font-medium">
              {lookupRate?.shortValue ?? '\u2014'}
            </div>
            <div className="text-xs text-ctp-overlay2">
              {lookupRate?.longValue ?? 'lookup rate'}
            </div>
          </div>
          {knownWordsSummary && knownWordsSummary.totalUniqueWords > 0 ? (
            <div>
              <div className="text-ctp-green font-medium">
                {formatNumber(knownWordsSummary.knownWordCount)} /{' '}
                {formatNumber(knownWordsSummary.totalUniqueWords)}
              </div>
              <div className="text-xs text-ctp-overlay2">
                known unique words (
                {Math.round(
                  (knownWordsSummary.knownWordCount / knownWordsSummary.totalUniqueWords) * 100,
                )}
                %)
              </div>
            </div>
          ) : (
            <div>
              <div className="text-ctp-peach font-medium">{formatPercent(knownTokenRate)}</div>
              <div className="text-xs text-ctp-overlay2">known token match rate</div>
            </div>
          )}
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
