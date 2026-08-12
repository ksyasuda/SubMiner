import { useId, useRef, useState } from 'react';
import { apiClient } from '../../lib/api-client';
import { formatDuration, formatNumber } from '../../lib/formatters';
import { useModalFocus } from '../../hooks/useModalFocus';
import { AnimeCoverImage } from './AnimeCoverImage';
import type { AnimeLibraryItem } from '../../types/stats';

interface AnimeMergeDialogProps {
  entries: AnimeLibraryItem[];
  onClose: () => void;
  onMerged: (survivingAnimeId: number) => void;
}

/** Biggest entry first: the one most likely to carry the right title and art. */
function pickDefaultKeeper(entries: AnimeLibraryItem[]): number {
  const best = [...entries].sort(
    (a, b) => b.episodeCount - a.episodeCount || b.totalActiveMs - a.totalActiveMs,
  )[0];
  return best?.animeId ?? 0;
}

export function AnimeMergeDialog({ entries, onClose, onMerged }: AnimeMergeDialogProps) {
  const headingId = useId();
  const dialogRef = useRef<HTMLDivElement>(null);
  const closeButtonRef = useRef<HTMLButtonElement>(null);
  const [keeperId, setKeeperId] = useState(() => pickDefaultKeeper(entries));
  const [merging, setMerging] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const totalEpisodes = entries.reduce((sum, entry) => sum + entry.episodeCount, 0);
  const totalCards = entries.reduce((sum, entry) => sum + entry.totalCards, 0);
  const totalActiveMs = entries.reduce((sum, entry) => sum + entry.totalActiveMs, 0);

  useModalFocus({
    dialogRef,
    initialFocusRef: closeButtonRef,
    dismissDisabled: merging,
    onDismiss: onClose,
  });

  const handleMerge = async () => {
    const sourceAnimeIds = entries
      .map((entry) => entry.animeId)
      .filter((animeId) => animeId !== keeperId);
    if (sourceAnimeIds.length === 0) return;
    setMerging(true);
    setError(null);
    try {
      const result = await apiClient.mergeAnime(keeperId, sourceAnimeIds);
      onMerged(result.animeId);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to merge these entries.');
      setMerging(false);
    }
  };

  // Dismissing mid-request would leave the caller unaware of a merge that is
  // still going to land, so the backdrop and close button are inert until it
  // resolves.
  const handleDismiss = () => {
    if (!merging) onClose();
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center pt-[10vh]"
      onClick={handleDismiss}
    >
      <div className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]" />
      <div
        ref={dialogRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby={headingId}
        className="relative bg-ctp-base border border-ctp-surface1 rounded-xl shadow-2xl w-full max-w-lg max-h-[70vh] flex flex-col animate-fade-in"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="p-4 border-b border-ctp-surface1">
          <div className="flex items-center justify-between">
            <h3 id={headingId} className="text-sm font-semibold text-ctp-text">
              Merge {entries.length} Library Entries
            </h3>
            <button
              ref={closeButtonRef}
              type="button"
              onClick={handleDismiss}
              disabled={merging}
              aria-label="Close"
              className="text-ctp-overlay2 hover:text-ctp-text text-lg leading-none disabled:opacity-50"
            >
              {'✕'}
            </button>
          </div>
          <p className="text-xs text-ctp-overlay2 mt-2">
            Pick the entry to keep. Every episode moves onto it and the others are removed; no
            sessions or mined cards are deleted.
          </p>
        </div>

        <div className="flex-1 overflow-y-auto p-2">
          {entries.map((entry) => (
            <button
              key={entry.animeId}
              type="button"
              disabled={merging}
              aria-pressed={keeperId === entry.animeId}
              onClick={() => setKeeperId(entry.animeId)}
              className={`w-full flex items-center gap-3 p-2.5 rounded-lg transition-colors text-left disabled:opacity-50 ${
                keeperId === entry.animeId ? 'bg-ctp-surface1' : 'hover:bg-ctp-surface0'
              }`}
            >
              <span
                aria-hidden="true"
                className={`w-4 h-4 rounded-full border shrink-0 ${
                  keeperId === entry.animeId
                    ? 'border-ctp-blue bg-ctp-blue'
                    : 'border-ctp-surface2 bg-transparent'
                }`}
              />
              <AnimeCoverImage
                animeId={entry.animeId}
                title={entry.canonicalTitle}
                coverRetryToken={entry.anilistId ?? 0}
                className="w-10 h-14 rounded shrink-0"
              />
              <div className="min-w-0 flex-1">
                <div className="text-sm text-ctp-text truncate">{entry.canonicalTitle}</div>
                <div className="text-xs text-ctp-overlay2 mt-0.5">
                  {entry.episodeCount} episode{entry.episodeCount !== 1 ? 's' : ''} ·{' '}
                  {formatDuration(entry.totalActiveMs)} · {formatNumber(entry.totalCards)} cards
                </div>
              </div>
              {keeperId === entry.animeId ? (
                <span className="text-xs text-ctp-blue shrink-0">Keep</span>
              ) : null}
            </button>
          ))}
        </div>

        <div className="p-4 border-t border-ctp-surface1 space-y-2">
          {error ? (
            <div role="alert" className="text-xs text-ctp-red">
              {error}
            </div>
          ) : null}
          <div className="flex items-center justify-between gap-3">
            <div className="text-xs text-ctp-overlay2">
              Result: {totalEpisodes} episode{totalEpisodes !== 1 ? 's' : ''} ·{' '}
              {formatDuration(totalActiveMs)} · {formatNumber(totalCards)} cards
            </div>
            <button
              type="button"
              disabled={merging}
              onClick={() => void handleMerge()}
              className="px-3 py-1.5 rounded-lg bg-ctp-blue/15 border border-ctp-blue/40 text-xs text-ctp-blue hover:bg-ctp-blue/25 transition-colors disabled:opacity-50"
            >
              {merging ? 'Merging…' : 'Merge Entries'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
