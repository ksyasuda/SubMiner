import { useState, useEffect, useRef } from 'react';
import { apiClient } from '../../lib/api-client';

interface AnilistMedia {
  id: number;
  episodes: number | null;
  season: string | null;
  seasonYear: number | null;
  description: string | null;
  coverImage: { large: string | null; medium: string | null } | null;
  title: { romaji: string | null; english: string | null; native: string | null } | null;
}

interface AnilistSelectorProps {
  animeId: number;
  initialQuery: string;
  onClose: () => void;
  onLinked: () => void;
}

export function AnilistSelector({
  animeId,
  initialQuery,
  onClose,
  onLinked,
}: AnilistSelectorProps) {
  const [query, setQuery] = useState(initialQuery);
  const [results, setResults] = useState<AnilistMedia[]>([]);
  const [loading, setLoading] = useState(false);
  const [linking, setLinking] = useState<number | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    inputRef.current?.focus();
    if (initialQuery) doSearch(initialQuery);
  }, []);

  const doSearch = async (q: string) => {
    if (!q.trim()) {
      setResults([]);
      return;
    }
    setLoading(true);
    try {
      const data = await apiClient.searchAnilist(q.trim());
      setResults(data);
    } catch {
      setResults([]);
    }
    setLoading(false);
  };

  const handleInput = (value: string) => {
    setQuery(value);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => doSearch(value), 400);
  };

  const handleSelect = async (media: AnilistMedia) => {
    setLinking(media.id);
    try {
      await apiClient.reassignAnimeAnilist(animeId, {
        anilistId: media.id,
        titleRomaji: media.title?.romaji ?? null,
        titleEnglish: media.title?.english ?? null,
        titleNative: media.title?.native ?? null,
        episodesTotal: media.episodes ?? null,
        description: media.description ?? null,
        coverUrl: media.coverImage?.large ?? media.coverImage?.medium ?? null,
      });
      onLinked();
    } catch {
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
            <h3 className="text-sm font-semibold text-ctp-text">Select AniList Entry</h3>
            <button
              type="button"
              onClick={onClose}
              className="text-ctp-overlay2 hover:text-ctp-text text-lg leading-none"
            >
              {'\u2715'}
            </button>
          </div>
          <input
            ref={inputRef}
            type="text"
            value={query}
            onChange={(e) => handleInput(e.target.value)}
            placeholder="Search AniList..."
            className="w-full bg-ctp-surface0 border border-ctp-surface1 rounded-lg px-3 py-2 text-sm text-ctp-text placeholder:text-ctp-overlay2 focus:outline-none focus:border-ctp-blue"
          />
        </div>

        <div className="flex-1 overflow-y-auto p-2">
          {loading && <div className="text-xs text-ctp-overlay2 p-3">Searching...</div>}
          {!loading && results.length === 0 && query.trim() && (
            <div className="text-xs text-ctp-overlay2 p-3">No results</div>
          )}
          {results.map((media) => (
            <button
              key={media.id}
              type="button"
              disabled={linking !== null}
              onClick={() => void handleSelect(media)}
              className="w-full flex items-center gap-3 p-2.5 rounded-lg hover:bg-ctp-surface0 transition-colors text-left disabled:opacity-50"
            >
              {media.coverImage?.medium ? (
                <img
                  src={media.coverImage.medium}
                  alt=""
                  className="w-10 h-14 rounded object-cover shrink-0 bg-ctp-surface1"
                />
              ) : (
                <div className="w-10 h-14 rounded bg-ctp-surface1 shrink-0" />
              )}
              <div className="min-w-0 flex-1">
                <div className="text-sm text-ctp-text truncate">
                  {media.title?.romaji ?? media.title?.english ?? 'Unknown'}
                </div>
                {media.title?.english && media.title.english !== media.title.romaji && (
                  <div className="text-xs text-ctp-subtext0 truncate">{media.title.english}</div>
                )}
                <div className="text-xs text-ctp-overlay2 mt-0.5">
                  {media.episodes ? `${media.episodes} eps` : 'Unknown eps'}
                  {media.seasonYear ? ` · ${media.season ?? ''} ${media.seasonYear}` : ''}
                </div>
              </div>
              {linking === media.id ? (
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
