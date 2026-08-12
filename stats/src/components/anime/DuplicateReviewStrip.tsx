import type { AnimeLibraryItem } from '../../types/stats';

interface DuplicateReviewStripProps {
  entries: AnimeLibraryItem[];
  current: number;
  total: number;
  dismissing: boolean;
  error?: string | null;
  onReview: () => void;
  onDismiss: () => void;
}

export function DuplicateReviewStrip({
  entries,
  current,
  total,
  dismissing,
  error = null,
  onReview,
  onDismiss,
}: DuplicateReviewStripProps) {
  return (
    <aside
      aria-label="Possible duplicate library entries"
      className="relative overflow-hidden rounded-lg border border-ctp-yellow/25 bg-ctp-yellow/[0.06] px-3 py-2.5"
    >
      <div className="absolute inset-y-0 left-0 w-0.5 bg-ctp-yellow/70" aria-hidden="true" />
      <div className="flex items-center gap-3">
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2">
            <span className="text-xs font-medium text-ctp-yellow">Possible duplicate</span>
            {total > 1 ? (
              <span className="text-[10px] tabular-nums text-ctp-overlay1">
                {current} of {total}
              </span>
            ) : null}
          </div>
          <p className="mt-0.5 truncate text-xs text-ctp-subtext0">
            {entries.map((entry) => entry.canonicalTitle).join(' · ')}
          </p>
        </div>
        <button
          type="button"
          disabled={dismissing}
          onClick={onDismiss}
          className="shrink-0 rounded-md px-2.5 py-1.5 text-xs text-ctp-overlay2 transition-colors hover:bg-ctp-surface0 hover:text-ctp-text disabled:opacity-50"
        >
          {dismissing ? 'Dismissing…' : 'Not duplicates'}
        </button>
        <button
          type="button"
          disabled={dismissing}
          onClick={onReview}
          className="shrink-0 rounded-md border border-ctp-yellow/35 bg-ctp-yellow/10 px-2.5 py-1.5 text-xs font-medium text-ctp-yellow transition-colors hover:bg-ctp-yellow/20 disabled:opacity-50"
        >
          Review merge
        </button>
      </div>
      {error ? (
        <p role="alert" className="mt-1.5 text-xs text-ctp-red">
          {error}
        </p>
      ) : null}
    </aside>
  );
}
