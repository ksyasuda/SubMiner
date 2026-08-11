import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { apiClient } from '../../lib/api-client';
import { formatDuration } from '../../lib/formatters';
import { AnimeCoverImage } from './AnimeCoverImage';
import type { AnimeLibraryItem } from '../../types/stats';

interface LibraryEntryPickerProps {
  heading: string;
  /** Entries that cannot be picked, typically the one being moved away from. */
  excludeAnimeIds?: number[];
  initialQuery?: string;
  busyAnimeId?: number | null;
  error?: string | null;
  onSelect: (entry: AnimeLibraryItem) => void;
  onClose: () => void;
}

export function LibraryEntryPicker({
  heading,
  excludeAnimeIds = [],
  initialQuery = '',
  busyAnimeId = null,
  error = null,
  onSelect,
  onClose,
}: LibraryEntryPickerProps) {
  const [entries, setEntries] = useState<AnimeLibraryItem[] | null>(null);
  const [loadFailed, setLoadFailed] = useState(false);
  const [query, setQuery] = useState(initialQuery);
  const inputRef = useRef<HTMLInputElement>(null);
  const headingId = useId();
  const searchId = useId();

  useEffect(() => {
    inputRef.current?.focus();
    let cancelled = false;
    apiClient
      .getAnimeLibrary()
      .then((data) => {
        if (!cancelled) setEntries(data);
      })
      .catch(() => {
        // Distinct from an empty library: telling the user "no other titles"
        // when the request failed hides a retryable error.
        if (cancelled) return;
        setEntries([]);
        setLoadFailed(true);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const excluded = useMemo(() => new Set(excludeAnimeIds), [excludeAnimeIds]);
  const visible = useMemo(() => {
    const term = query.trim().toLowerCase();
    return (entries ?? [])
      .filter((entry) => !excluded.has(entry.animeId))
      .filter((entry) => !term || entry.canonicalTitle.toLowerCase().includes(term))
      .sort((a, b) => b.lastWatchedMs - a.lastWatchedMs);
  }, [entries, excluded, query]);

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center pt-[10vh]" onClick={onClose}>
      <div className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]" />
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby={headingId}
        className="relative bg-ctp-base border border-ctp-surface1 rounded-xl shadow-2xl w-full max-w-lg max-h-[70vh] flex flex-col animate-fade-in"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="p-4 border-b border-ctp-surface1">
          <div className="flex items-center justify-between mb-3">
            <h3 id={headingId} className="text-sm font-semibold text-ctp-text">
              {heading}
            </h3>
            <button
              type="button"
              onClick={onClose}
              aria-label="Close"
              className="text-ctp-overlay2 hover:text-ctp-text text-lg leading-none"
            >
              {'✕'}
            </button>
          </div>
          <label htmlFor={searchId} className="sr-only">
            Search library
          </label>
          <input
            ref={inputRef}
            id={searchId}
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search library..."
            className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2 text-sm text-ctp-text placeholder:text-ctp-overlay2 focus:outline-none focus:border-ctp-blue"
          />
          {error ? (
            <div role="alert" className="text-xs text-ctp-red mt-2">
              {error}
            </div>
          ) : null}
        </div>

        <div className="flex-1 overflow-y-auto p-2">
          {entries === null && <div className="text-xs text-ctp-overlay2 p-3">Loading...</div>}
          {loadFailed && (
            <div role="alert" className="text-xs text-ctp-red p-3">
              Could not load the library. Close this dialog and try again.
            </div>
          )}
          {!loadFailed && entries !== null && visible.length === 0 && (
            <div className="text-xs text-ctp-overlay2 p-3">
              {query.trim() ? 'No matches' : 'No other titles'}
            </div>
          )}
          {visible.map((entry) => (
            <button
              key={entry.animeId}
              type="button"
              disabled={busyAnimeId !== null}
              onClick={() => onSelect(entry)}
              className="w-full flex items-center gap-3 p-2.5 rounded-lg hover:bg-ctp-surface0 transition-colors text-left disabled:opacity-50"
            >
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
                  {formatDuration(entry.totalActiveMs)}
                </div>
              </div>
              {busyAnimeId === entry.animeId ? (
                <span className="text-xs text-ctp-blue shrink-0">Moving...</span>
              ) : (
                <span className="text-xs text-ctp-overlay2 shrink-0">Select</span>
              )}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}
