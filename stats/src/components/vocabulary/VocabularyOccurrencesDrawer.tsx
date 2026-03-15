import type { KanjiEntry, VocabularyEntry, VocabularyOccurrenceEntry } from '../../types/stats';
import { formatNumber } from '../../lib/formatters';

type VocabularyDrawerTarget =
  | {
      kind: 'word';
      entry: VocabularyEntry;
    }
  | {
      kind: 'kanji';
      entry: KanjiEntry;
    };

interface VocabularyOccurrencesDrawerProps {
  target: VocabularyDrawerTarget | null;
  occurrences: VocabularyOccurrenceEntry[];
  loading: boolean;
  loadingMore: boolean;
  error: string | null;
  hasMore: boolean;
  onClose: () => void;
  onLoadMore: () => void;
}

function formatSegment(ms: number | null): string {
  if (ms == null || !Number.isFinite(ms)) return '--:--';
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${String(seconds).padStart(2, '0')}`;
}

function renderTitle(target: VocabularyDrawerTarget): string {
  return target.kind === 'word' ? target.entry.headword : target.entry.kanji;
}

function renderSubtitle(target: VocabularyDrawerTarget): string {
  if (target.kind === 'word') {
    return target.entry.reading || target.entry.word;
  }
  return `${formatNumber(target.entry.frequency)} seen`;
}

function renderFrequency(target: VocabularyDrawerTarget): string {
  return `${formatNumber(target.entry.frequency)} total`;
}

export function VocabularyOccurrencesDrawer({
  target,
  occurrences,
  loading,
  loadingMore,
  error,
  hasMore,
  onClose,
  onLoadMore,
}: VocabularyOccurrencesDrawerProps) {
  if (!target) return null;

  return (
    <div className="fixed inset-0 z-40">
      <button
        type="button"
        aria-label="Close occurrence drawer"
        className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]"
        onClick={onClose}
      />
      <aside className="absolute right-0 top-0 h-full w-full max-w-xl border-l border-ctp-surface1 bg-ctp-mantle shadow-2xl">
        <div className="flex h-full flex-col">
          <div className="flex items-start justify-between border-b border-ctp-surface1 px-5 py-4">
            <div className="min-w-0">
              <div className="text-xs uppercase tracking-[0.18em] text-ctp-overlay1">
                {target.kind === 'word' ? 'Word Occurrences' : 'Kanji Occurrences'}
              </div>
              <h2 className="mt-1 truncate text-2xl font-semibold text-ctp-text">
                {renderTitle(target)}
              </h2>
              <div className="mt-1 text-sm text-ctp-subtext0">{renderSubtitle(target)}</div>
              <div className="mt-2 text-xs text-ctp-overlay1">
                {renderFrequency(target)} · {formatNumber(occurrences.length)} loaded
              </div>
            </div>
            <button
              type="button"
              className="rounded-md border border-ctp-surface2 px-3 py-1.5 text-xs font-medium text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue"
              onClick={onClose}
            >
              Close
            </button>
          </div>

          <div className="flex-1 overflow-y-auto px-4 py-4">
            {loading ? <div className="text-sm text-ctp-overlay2">Loading occurrences...</div> : null}
            {!loading && error ? <div className="text-sm text-ctp-red">Error: {error}</div> : null}
            {!loading && !error && occurrences.length === 0 ? (
              <div className="text-sm text-ctp-overlay2">No occurrences tracked yet.</div>
            ) : null}
            {!loading && !error ? (
              <div className="space-y-3">
                {occurrences.map((occurrence, index) => (
                  <article
                    key={`${occurrence.sessionId}-${occurrence.lineIndex}-${occurrence.segmentStartMs ?? index}`}
                    className="rounded-xl border border-ctp-surface1 bg-ctp-surface0/90 p-4"
                  >
                    <div className="flex items-start justify-between gap-3">
                      <div className="min-w-0">
                        <div className="truncate text-sm font-semibold text-ctp-text">
                          {occurrence.animeTitle ?? occurrence.videoTitle}
                        </div>
                        <div className="truncate text-xs text-ctp-subtext0">
                          {occurrence.videoTitle} · line {occurrence.lineIndex}
                        </div>
                      </div>
                      <div className="rounded-full bg-ctp-blue/10 px-2 py-1 text-[11px] font-medium text-ctp-blue">
                        {formatNumber(occurrence.occurrenceCount)} in line
                      </div>
                    </div>
                    <div className="mt-3 text-xs text-ctp-overlay1">
                      {formatSegment(occurrence.segmentStartMs)}-{formatSegment(occurrence.segmentEndMs)} · session{' '}
                      {occurrence.sessionId}
                    </div>
                    <p className="mt-3 rounded-lg bg-ctp-base/70 px-3 py-3 text-sm leading-6 text-ctp-text">
                      {occurrence.text}
                    </p>
                  </article>
                ))}
              </div>
            ) : null}
          </div>

          {!loading && !error && hasMore ? (
            <div className="border-t border-ctp-surface1 px-4 py-4">
              <button
                type="button"
                className="w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-blue hover:text-ctp-blue disabled:cursor-not-allowed disabled:opacity-60"
                onClick={onLoadMore}
                disabled={loadingMore}
              >
                {loadingMore ? 'Loading more...' : 'Load more'}
              </button>
            </div>
          ) : null}
        </div>
      </aside>
    </div>
  );
}

export type { VocabularyDrawerTarget };
