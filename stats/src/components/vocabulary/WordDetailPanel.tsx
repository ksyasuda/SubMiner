import { useRef, useState } from 'react';
import { useWordDetail } from '../../hooks/useWordDetail';
import { apiClient } from '../../lib/api-client';
import { formatNumber, formatRelativeDate } from '../../lib/formatters';
import type { VocabularyOccurrenceEntry } from '../../types/stats';
import { PosBadge } from './pos-helpers';

const OCCURRENCES_PAGE_SIZE = 50;

interface WordDetailPanelProps {
  wordId: number | null;
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

export function WordDetailPanel({ wordId, onClose, onSelectWord, onNavigateToAnime }: WordDetailPanelProps) {
  const { data, loading, error } = useWordDetail(wordId);
  const [occurrences, setOccurrences] = useState<VocabularyOccurrenceEntry[]>([]);
  const [occLoading, setOccLoading] = useState(false);
  const [occLoadingMore, setOccLoadingMore] = useState(false);
  const [occError, setOccError] = useState<string | null>(null);
  const [hasMore, setHasMore] = useState(false);
  const [occLoaded, setOccLoaded] = useState(false);
  const requestIdRef = useRef(0);

  if (wordId === null) return null;

  const loadOccurrences = async (detail: NonNullable<typeof data>['detail'], offset: number, append: boolean) => {
    const reqId = ++requestIdRef.current;
    if (append) {
      setOccLoadingMore(true);
    } else {
      setOccLoading(true);
      setOccError(null);
    }
    try {
      const rows = await apiClient.getWordOccurrences(
        detail.headword, detail.word, detail.reading,
        OCCURRENCES_PAGE_SIZE, offset,
      );
      if (reqId !== requestIdRef.current) return;
      setOccurrences(prev => append ? [...prev, ...rows] : rows);
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
    void loadOccurrences(data.detail, 0, false);
  };

  const handleLoadMore = () => {
    if (!data || occLoadingMore || !hasMore) return;
    void loadOccurrences(data.detail, occurrences.length, true);
  };

  return (
    <div className="fixed inset-0 z-40">
      <button
        type="button"
        aria-label="Close word detail panel"
        className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]"
        onClick={onClose}
      />
      <aside className="absolute right-0 top-0 h-full w-full max-w-xl border-l border-ctp-surface1 bg-ctp-mantle shadow-2xl">
        <div className="flex h-full flex-col">
          <div className="flex items-start justify-between border-b border-ctp-surface1 px-5 py-4">
            <div className="min-w-0">
              <div className="text-xs uppercase tracking-[0.18em] text-ctp-overlay1">Word Detail</div>
              {loading && <div className="mt-2 text-sm text-ctp-overlay2">Loading...</div>}
              {error && <div className="mt-2 text-sm text-ctp-red">Error: {error}</div>}
              {data && (
                <>
                  <h2 className="mt-1 truncate text-3xl font-semibold text-ctp-text">{data.detail.headword}</h2>
                  <div className="mt-1 text-sm text-ctp-subtext0">{data.detail.reading || data.detail.word}</div>
                  <div className="mt-2 flex flex-wrap gap-1.5">
                    {data.detail.partOfSpeech && <PosBadge pos={data.detail.partOfSpeech} />}
                    {data.detail.pos1 && data.detail.pos1 !== data.detail.partOfSpeech && (
                      <span className="rounded-full bg-ctp-surface1 px-2 py-0.5 text-[11px] text-ctp-subtext0">{data.detail.pos1}</span>
                    )}
                    {data.detail.pos2 && (
                      <span className="rounded-full bg-ctp-surface1 px-2 py-0.5 text-[11px] text-ctp-subtext0">{data.detail.pos2}</span>
                    )}
                    {data.detail.pos3 && (
                      <span className="rounded-full bg-ctp-surface1 px-2 py-0.5 text-[11px] text-ctp-subtext0">{data.detail.pos3}</span>
                    )}
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
                    <div className="text-lg font-bold text-ctp-blue">{formatNumber(data.detail.frequency)}</div>
                    <div className="text-[11px] text-ctp-overlay1 uppercase">Frequency</div>
                  </div>
                  <div className="rounded-lg bg-ctp-surface0 p-3 text-center">
                    <div className="text-sm font-medium text-ctp-green">{formatRelativeDate(data.detail.firstSeen)}</div>
                    <div className="text-[11px] text-ctp-overlay1 uppercase">First Seen</div>
                  </div>
                  <div className="rounded-lg bg-ctp-surface0 p-3 text-center">
                    <div className="text-sm font-medium text-ctp-mauve">{formatRelativeDate(data.detail.lastSeen)}</div>
                    <div className="text-[11px] text-ctp-overlay1 uppercase">Last Seen</div>
                  </div>
                </div>

                {data.animeAppearances.length > 0 && (
                  <section>
                    <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">Anime Appearances</h3>
                    <div className="space-y-1.5">
                      {data.animeAppearances.map(a => (
                        <button
                          key={a.animeId}
                          type="button"
                          onClick={() => { onClose(); onNavigateToAnime?.(a.animeId); }}
                          className="w-full flex items-center justify-between rounded-lg bg-ctp-surface0 px-3 py-2 text-sm transition hover:border-ctp-blue hover:ring-1 hover:ring-ctp-blue text-left"
                        >
                          <span className="truncate text-ctp-text">{a.animeTitle}</span>
                          <span className="ml-2 shrink-0 rounded-full bg-ctp-blue/10 px-2 py-0.5 text-[11px] font-medium text-ctp-blue">
                            {formatNumber(a.occurrenceCount)}
                          </span>
                        </button>
                      ))}
                    </div>
                  </section>
                )}

                {data.similarWords.length > 0 && (
                  <section>
                    <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">Similar Words</h3>
                    <div className="flex flex-wrap gap-1.5">
                      {data.similarWords.map(sw => (
                        <button
                          key={sw.wordId}
                          type="button"
                          className="inline-flex items-center gap-1 rounded px-2 py-1 text-xs text-ctp-blue bg-ctp-blue/10 transition hover:ring-1 hover:ring-ctp-blue"
                          onClick={() => onSelectWord?.(sw.wordId)}
                        >
                          {sw.headword}
                          <span className="opacity-60">({formatNumber(sw.frequency)})</span>
                        </button>
                      ))}
                    </div>
                  </section>
                )}

                <section>
                  <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">Example Lines</h3>
                  {!occLoaded && !occLoading && (
                    <button
                      type="button"
                      className="w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-blue hover:text-ctp-blue"
                      onClick={handleShowOccurrences}
                    >
                      Load example lines
                    </button>
                  )}
                  {occLoading && <div className="text-sm text-ctp-overlay2">Loading occurrences...</div>}
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
                            <div className="rounded-full bg-ctp-blue/10 px-2 py-1 text-[11px] font-medium text-ctp-blue">
                              {formatNumber(occ.occurrenceCount)} in line
                            </div>
                          </div>
                          <div className="mt-3 text-xs text-ctp-overlay1">
                            {formatSegment(occ.segmentStartMs)}-{formatSegment(occ.segmentEndMs)} · session {occ.sessionId}
                          </div>
                          <p className="mt-3 rounded-lg bg-ctp-base/70 px-3 py-3 text-sm leading-6 text-ctp-text">
                            {occ.text}
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
                className="w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-blue hover:text-ctp-blue disabled:cursor-not-allowed disabled:opacity-60"
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
