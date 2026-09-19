import { useState, useEffect, useRef } from 'react';
import { apiClient } from '../../lib/api-client';
import { normalizeAnilistSearchQuery } from '../../lib/anilist-search-query';
import type { StatsTmdbSearchResult } from '../../types/stats';

interface TmdbSelectorProps {
  animeId: number;
  initialQuery: string;
  onClose: () => void;
  onLinked: () => void;
}

// The stats API answers a missing key with 503 and the message from config.
function describeSearchError(err: unknown): string {
  const message = err instanceof Error ? err.message : '';
  if (/\b503\b/.test(message)) {
    return 'TMDB API key not configured. Set tmdb.apiKey or tmdb.apiKeyCommand in your config.';
  }
  return 'TMDB search failed. Check your connection and API key.';
}

export function TmdbSelector({ animeId, initialQuery, onClose, onLinked }: TmdbSelectorProps) {
  const [query, setQuery] = useState(() => normalizeAnilistSearchQuery(initialQuery));
  const [results, setResults] = useState<StatsTmdbSearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [linking, setLinking] = useState<number | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const searchSequenceRef = useRef(0);

  useEffect(() => {
    inputRef.current?.focus();
    const normalizedInitialQuery = normalizeAnilistSearchQuery(initialQuery);
    setQuery(normalizedInitialQuery);
    setResults([]);
    setError(null);
    setLoading(false);
    setLinking(null);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (normalizedInitialQuery) void doSearch(normalizedInitialQuery);
    return () => {
      searchSequenceRef.current += 1;
      if (debounceRef.current) clearTimeout(debounceRef.current);
    };
  }, [initialQuery, animeId]);

  const doSearch = async (q: string) => {
    const sequence = ++searchSequenceRef.current;
    const searchQuery = normalizeAnilistSearchQuery(q);
    if (!searchQuery) {
      setResults([]);
      setError(null);
      setLoading(false);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const nextResults = await apiClient.searchTmdb(searchQuery);
      if (sequence === searchSequenceRef.current) setResults(nextResults);
    } catch (err) {
      if (sequence !== searchSequenceRef.current) return;
      setResults([]);
      setError(describeSearchError(err));
    } finally {
      if (sequence === searchSequenceRef.current) setLoading(false);
    }
  };

  const handleInput = (value: string) => {
    searchSequenceRef.current += 1;
    setQuery(value);
    setResults([]);
    setError(null);
    const hasQuery = Boolean(normalizeAnilistSearchQuery(value));
    setLoading(hasQuery);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (!hasQuery) return;
    debounceRef.current = setTimeout(() => void doSearch(value), 400);
  };

  const handleSelect = async (media: StatsTmdbSearchResult) => {
    setLinking(media.tmdbId);
    setError(null);
    try {
      await apiClient.reassignAnimeTmdb(animeId, {
        tmdbId: media.tmdbId,
        tmdbType: media.tmdbType,
      });
      onLinked();
    } catch (err) {
      setError(describeSearchError(err));
      setLinking(null);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center pt-[10vh]" onClick={onClose}>
      <div className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]" />
      <div
        className="relative bg-ctp-base border border-ctp-surface1 rounded-xl shadow-2xl w-full max-w-lg max-h-[70vh] flex flex-col animate-fade-in"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="p-4 border-b border-ctp-surface1">
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-semibold text-ctp-text">Select TMDB Title</h3>
            <button
              type="button"
              onClick={onClose}
              className="text-ctp-overlay2 hover:text-ctp-text text-lg leading-none"
            >
              {'✕'}
            </button>
          </div>
          <input
            ref={inputRef}
            type="text"
            value={query}
            onChange={(e) => handleInput(e.target.value)}
            placeholder="Search TMDB for a drama or movie..."
            className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2 text-sm text-ctp-text placeholder:text-ctp-overlay2 focus:outline-none focus:border-ctp-blue"
          />
        </div>

        <div className="flex-1 overflow-y-auto p-2">
          {loading && <div className="text-xs text-ctp-overlay2 p-3">Searching...</div>}
          {error && <div className="text-xs text-ctp-red p-3">{error}</div>}
          {!loading && !error && results.length === 0 && query.trim() && (
            <div className="text-xs text-ctp-overlay2 p-3">No results</div>
          )}
          {results.map((media) => (
            <button
              key={`${media.tmdbType}-${media.tmdbId}`}
              type="button"
              disabled={linking !== null}
              onClick={() => void handleSelect(media)}
              className="w-full flex items-center gap-3 p-2.5 rounded-lg hover:bg-ctp-surface0 transition-colors text-left disabled:opacity-50"
            >
              {media.posterUrl ? (
                <img
                  src={media.posterUrl}
                  alt=""
                  className="w-10 h-14 rounded object-cover shrink-0 bg-ctp-surface1"
                />
              ) : (
                <div className="w-10 h-14 rounded bg-ctp-surface1 shrink-0" />
              )}
              <div className="min-w-0 flex-1">
                <div className="text-sm text-ctp-text truncate">{media.title}</div>
                {media.originalTitle !== media.title && (
                  <div className="text-xs text-ctp-subtext0 truncate">{media.originalTitle}</div>
                )}
                <div className="text-xs text-ctp-overlay2 mt-0.5">
                  {media.tmdbType === 'movie' ? 'Movie' : 'TV'}
                  {media.year ? ` · ${media.year}` : ''}
                  {media.isAnimation ? ' · Animation' : ''}
                </div>
              </div>
              {linking === media.tmdbId ? (
                <span className="text-xs text-ctp-blue shrink-0">Linking...</span>
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
