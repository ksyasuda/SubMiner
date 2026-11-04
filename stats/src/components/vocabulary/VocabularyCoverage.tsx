import { formatNumber } from '../../lib/formatters';
import type { VocabularyCoverageBand, VocabularyCoverageBandId } from '../../types/stats';

const BAND_LABELS: Record<VocabularyCoverageBandId, string> = {
  top1k: 'Top 1k',
  top5k: '1k to 5k',
  top10k: '5k to 10k',
  top20k: '10k to 20k',
};

/** How much of each frequency band you know, which points at where to mine next. */
export function VocabularyCoverage({ bands }: { bands: VocabularyCoverageBand[] }) {
  return (
    <div className="h-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
      <h3 className="text-xs font-semibold text-ctp-text">Known Words by Frequency</h3>
      <p className="mb-3 mt-0.5 text-[11px] text-ctp-overlay2">
        Words you have seen in each band that are already in Anki.
      </p>
      <div className="space-y-3">
        {bands.map(({ band, seen, known }) => {
          const percent = seen > 0 ? Math.round((known / seen) * 100) : 0;
          return (
            <div
              key={band}
              className="grid grid-cols-[4.5rem_minmax(0,1fr)_auto] items-center gap-2 text-xs"
            >
              <span className="text-ctp-subtext0">{BAND_LABELS[band]}</span>
              <div
                role="img"
                aria-label={`${percent}% known`}
                className="h-2 overflow-hidden rounded-full bg-ctp-surface1/60"
              >
                <div
                  className="h-full rounded-full bg-ctp-blue"
                  style={{ width: `${seen > 0 ? (known / seen) * 100 : 0}%` }}
                />
              </div>
              <span className="whitespace-nowrap text-right text-[11px] tabular-nums text-ctp-overlay2">
                {formatNumber(known)} of {formatNumber(seen)}, {percent}%
              </span>
            </div>
          );
        })}
      </div>
    </div>
  );
}
