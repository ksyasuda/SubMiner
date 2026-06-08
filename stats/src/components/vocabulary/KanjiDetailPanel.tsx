import { useRef, useState, useEffect } from 'react';
import { useKanjiDetail } from '../../hooks/useKanjiDetail';
import { apiClient } from '../../lib/api-client';
import { epochMsFromDbTimestamp, formatNumber, formatRelativeDate } from '../../lib/formatters';
import type { VocabularyOccurrenceEntry } from '../../types/stats';

const OCCURRENCES_PAGE_SIZE = 50;
const ANIME_APPEARANCES_LIMIT = 5;

interface KanjiDetailPanelProps {
  kanjiId: number | null;
  onClose: () => void;
  onSelectWord?: (wordId: number) => void;
  onNavigateToAnime?: (animeId: number) => void;
}

function formatSegment(ms: number | null): string {
  if (ms == null || !Number.isFinite(ms)) return '--:--';
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${String(seconds).padStart(2, '0')}`;
}

function highlightKanji(text: string, kanji: string) {
  if (!kanji) return text;
  const parts = text.split(kanji);
  if (parts.length === 1) return text;
  return parts.flatMap((part, idx) =>
    idx === 0
      ? [part]
      : [
          <mark
            key={idx}
            className="rounded bg-ctp-teal/20 px-0.5 font-semibold text-ctp-teal"
          >
            {kanji}
          </mark>,
          part,
        ],
  );
}

export function KanjiDetailPanel({
  kanjiId,
  onClose,
  onSelectWord,
  onNavigateToAnime,
}: KanjiDetailPanelProps) {
  const { data, loading, error } = useKanjiDetail(kanjiId);
  const [occurrences, setOccurrences] = useState<VocabularyOccurrenceEntry[]>([]);
  const [occLoading, setOccLoading] = useState(false);
  const [occLoadingMore, setOccLoadingMore] = useState(false);
  const [occError, setOccError] = useState<string | null>(null);
  const [hasMore, setHasMore] = useState(false);
  const [occLoaded, setOccLoaded] = useState(false);
  const [showAllAnime, setShowAllAnime] = useState(false);
  const requestIdRef = useRef(0);

  useEffect(() => {
    setOccurrences([]);
    setOccLoaded(false);
    setOccLoading(false);
    setOccLoadingMore(false);
    setOccError(null);
    setHasMore(false);
    setShowAllAnime(false);
    requestIdRef.current++;
  }, [kanjiId]);

  if (kanjiId === null) return null;

  const loadOccurrences = async (kanji: string, offset: number, append: boolean) => {
    const reqId = ++requestIdRef.current;
    if (append) {
      setOccLoadingMore(true);
    } else {
      setOccLoading(true);
      setOccError(null);
    }
    try {
      const rows = await apiClient.getKanjiOccurrences(kanji, OCCURRENCES_PAGE_SIZE, offset);
      if (reqId !== requestIdRef.current) return;
      setOccurrences((prev) => (append ? [...prev, ...rows] : rows));
      setHasMore(rows.length === OCCURRENCES_PAGE_SIZE);
    } catch (err) {
      if (reqId !== requestIdRef.current) return;
      setOccError(err instanceof Error ? err.message : String(err));
      if (!append) {
        setOccurrences([]);
        setHasMore(false);
      }
    } finally {
      if (reqId !== requestIdRef.current) return;
      setOccLoading(false);
      setOccLoadingMore(false);
      setOccLoaded(true);
    }
  };

  const handleShowOccurrences = () => {
    if (!data) return;
    void loadOccurrences(data.detail.kanji, 0, false);
  };

  const handleLoadMore = () => {
    if (!data || occLoadingMore || !hasMore) return;
    void loadOccurrences(data.detail.kanji, occurrences.length, true);
  };

  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center p-4">
      <button
        type="button"
        aria-label="Close kanji detail panel"
        className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]"
        onClick={onClose}
      />
      <aside className="relative flex max-h-[85vh] w-full max-w-2xl flex-col overflow-hidden rounded-xl border border-ctp-surface1 bg-ctp-mantle shadow-2xl">
        <div className="flex min-h-0 flex-1 flex-col">
          <div className="flex items-start justify-between border-b border-ctp-surface1 px-5 py-4">
            <div className="min-w-0">
              <div className="text-xs uppercase tracking-[0.18em] text-ctp-overlay1">
                Kanji Detail
              </div>
              {loading && <div className="mt-2 text-sm text-ctp-overlay2">Loading...</div>}
              {error && <div className="mt-2 text-sm text-ctp-red">Error: {error}</div>}
              {data && (
                <>
                  <h2 className="mt-1 text-5xl font-semibold text-ctp-teal">{data.detail.kanji}</h2>
                  <div className="mt-2 text-sm text-ctp-subtext0">
                    {formatNumber(data.detail.frequency)} total occurrences
                  </div>
                </>
              )}
            </div>
            <button
              type="button"
              className="rounded-md border border-ctp-surface2 px-3 py-1.5 text-xs font-medium text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue"
              onClick={onClose}
            >
              Close
            </button>
          </div>

          <div className="flex-1 overflow-y-auto px-5 py-4 space-y-5">
            {data && (
              <>
                <div className="grid grid-cols-3 gap-3">
                  <div className="rounded-lg bg-ctp-surface0 p-3 text-center">
                    <div className="text-lg font-bold text-ctp-teal">
                      {formatNumber(data.detail.frequency)}
                    </div>
                    <div className="text-[11px] text-ctp-overlay1 uppercase">Frequency</div>
                  </div>
                  <div className="rounded-lg bg-ctp-surface0 p-3 text-center">
                    <div className="text-sm font-medium text-ctp-green">
                      {formatRelativeDate(epochMsFromDbTimestamp(data.detail.firstSeen))}
                    </div>
                    <div className="text-[11px] text-ctp-overlay1 uppercase">First Seen</div>
                  </div>
                  <div className="rounded-lg bg-ctp-surface0 p-3 text-center">
                    <div className="text-sm font-medium text-ctp-mauve">
                      {formatRelativeDate(epochMsFromDbTimestamp(data.detail.lastSeen))}
                    </div>
                    <div className="text-[11px] text-ctp-overlay1 uppercase">Last Seen</div>
                  </div>
                </div>

                {data.animeAppearances.length > 0 && (
                  <section>
                    <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">
                      Anime Appearances
                    </h3>
                    <div className="space-y-1.5">
                      {(showAllAnime
                        ? data.animeAppearances
                        : data.animeAppearances.slice(0, ANIME_APPEARANCES_LIMIT)
                      ).map((a) => (
                        <button
                          key={a.animeId}
                          type="button"
                          onClick={() => {
                            onClose();
                            onNavigateToAnime?.(a.animeId);
                          }}
                          className="w-full flex items-center justify-between rounded-lg bg-ctp-surface0 px-3 py-2 text-sm transition hover:border-ctp-teal hover:ring-1 hover:ring-ctp-teal text-left"
                        >
                          <span className="truncate text-ctp-text">{a.animeTitle}</span>
                          <span className="ml-2 shrink-0 rounded-full bg-ctp-teal/10 px-2 py-0.5 text-[11px] font-medium text-ctp-teal">
                            {formatNumber(a.occurrenceCount)}
                          </span>
                        </button>
                      ))}
                    </div>
                    {data.animeAppearances.length > ANIME_APPEARANCES_LIMIT && (
                      <button
                        type="button"
                        className="mt-2 w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-teal hover:text-ctp-teal"
                        onClick={() => setShowAllAnime((prev) => !prev)}
                      >
                        {showAllAnime
                          ? 'Show less'
                          : `Show ${formatNumber(
                              data.animeAppearances.length - ANIME_APPEARANCES_LIMIT,
                            )} more`}
                      </button>
                    )}
                  </section>
                )}

                {data.words.length > 0 && (
                  <section>
                    <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">
                      Words Using This Kanji
                    </h3>
                    <div className="flex flex-wrap gap-1.5">
                      {data.words.map((w) => (
                        <button
                          key={w.wordId}
                          type="button"
                          className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-ctp-blue bg-ctp-blue/10 transition hover:ring-1 hover:ring-ctp-blue"
                          onClick={() => onSelectWord?.(w.wordId)}
                        >
                          {w.headword}
                          <span className="opacity-60">({formatNumber(w.frequency)})</span>
                        </button>
                      ))}
                    </div>
                  </section>
                )}

                <section>
                  <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">
                    Example Lines
                  </h3>
                  {!occLoaded && !occLoading && (
                    <button
                      type="button"
                      className="w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-teal hover:text-ctp-teal"
                      onClick={handleShowOccurrences}
                    >
                      Load example lines
                    </button>
                  )}
                  {occLoading && (
                    <div className="text-sm text-ctp-overlay2">Loading occurrences...</div>
                  )}
                  {occError && <div className="text-sm text-ctp-red">Error: {occError}</div>}
                  {occLoaded && !occLoading && occurrences.length === 0 && (
                    <div className="text-sm text-ctp-overlay2">No occurrences tracked yet.</div>
                  )}
                  {occurrences.length > 0 && (
                    <div className="space-y-3">
                      {occurrences.map((occ, idx) => (
                        <article
                          key={`${occ.sessionId}-${occ.lineIndex}-${occ.segmentStartMs ?? idx}`}
                          className="rounded-xl border border-ctp-surface1 bg-ctp-surface0/90 p-4"
                        >
                          <div className="flex items-start justify-between gap-3">
                            <div className="min-w-0">
                              <div className="truncate text-sm font-semibold text-ctp-text">
                                {occ.animeTitle ?? occ.videoTitle}
                              </div>
                              <div className="truncate text-xs text-ctp-subtext0">
                                {occ.videoTitle} · line {occ.lineIndex}
                              </div>
                            </div>
                            <div className="rounded-full bg-ctp-teal/10 px-2 py-1 text-[11px] font-medium text-ctp-teal">
                              {formatNumber(occ.occurrenceCount)} in line
                            </div>
                          </div>
                          <div className="mt-3 text-xs text-ctp-overlay1">
                            {formatSegment(occ.segmentStartMs)}-{formatSegment(occ.segmentEndMs)} ·
                            session {occ.sessionId}
                          </div>
                          <p className="mt-3 rounded-lg bg-ctp-base/70 px-3 py-3 text-sm leading-6 text-ctp-text">
                            {highlightKanji(occ.text, data.detail.kanji)}
                          </p>
                        </article>
                      ))}
                    </div>
                  )}
                </section>
              </>
            )}
          </div>

          {occLoaded && !occLoading && !occError && hasMore && (
            <div className="border-t border-ctp-surface1 px-4 py-4">
              <button
                type="button"
                className="w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-teal hover:text-ctp-teal disabled:cursor-not-allowed disabled:opacity-60"
                onClick={handleLoadMore}
                disabled={occLoadingMore}
              >
                {occLoadingMore ? 'Loading more...' : 'Load more'}
              </button>
            </div>
          )}
        </div>
      </aside>
    </div>
  );
}
