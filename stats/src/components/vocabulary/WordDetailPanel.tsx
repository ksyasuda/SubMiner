import { useRef, useState, useEffect } from 'react';
import { useWordDetail } from '../../hooks/useWordDetail';
import { apiClient } from '../../lib/api-client';
import { epochMsFromDbTimestamp, formatNumber, formatRelativeDate } from '../../lib/formatters';
import {
  buildStatsMineCardParams,
  getStatsMineCardError,
  getStatsMineCardUnavailableReason,
} from '../../lib/mining';
import { fullReading } from '../../lib/reading-utils';
import type { VocabularyOccurrenceEntry } from '../../types/stats';
import { PosBadge } from './pos-helpers';

const INITIAL_PAGE_SIZE = 5;
const LOAD_MORE_SIZE = 10;
const MEDIA_APPEARANCES_LIMIT = 5;

type MineStatus = { loading?: boolean; success?: boolean; error?: string };

interface WordDetailPanelProps {
  wordId: number | null;
  onClose: () => void;
  onSelectWord?: (wordId: number) => void;
  onNavigateToAnime?: (animeId: number) => void;
  isExcluded?: (w: { headword: string; word: string; reading: string }) => boolean;
  onToggleExclusion?: (w: { headword: string; word: string; reading: string }) => void;
}

function highlightWord(text: string, words: string[]): React.ReactNode {
  const needles = words.filter(Boolean);
  if (needles.length === 0) return text;

  const escaped = needles.map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  const pattern = new RegExp(`(${escaped.join('|')})`, 'g');
  const parts = text.split(pattern);
  const needleSet = new Set(needles);

  return parts.map((part, i) =>
    needleSet.has(part) ? (
      <mark
        key={i}
        className="bg-transparent text-ctp-blue underline decoration-ctp-blue/40 underline-offset-2"
      >
        {part}
      </mark>
    ) : (
      part
    ),
  );
}

function formatSegment(ms: number | null): string {
  if (ms == null || !Number.isFinite(ms)) return '--:--';
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${String(seconds).padStart(2, '0')}`;
}

export function WordDetailPanel({
  wordId,
  onClose,
  onSelectWord,
  onNavigateToAnime,
  isExcluded,
  onToggleExclusion,
}: WordDetailPanelProps) {
  const { data, loading, error } = useWordDetail(wordId);
  const [occurrences, setOccurrences] = useState<VocabularyOccurrenceEntry[]>([]);
  const [occLoading, setOccLoading] = useState(false);
  const [occLoadingMore, setOccLoadingMore] = useState(false);
  const [occError, setOccError] = useState<string | null>(null);
  const [hasMore, setHasMore] = useState(false);
  const [occLoaded, setOccLoaded] = useState(false);
  const [mineStatus, setMineStatus] = useState<Record<string, MineStatus>>({});
  const [showAllAnime, setShowAllAnime] = useState(false);
  const requestIdRef = useRef(0);

  useEffect(() => {
    setOccurrences([]);
    setOccLoaded(false);
    setOccLoading(false);
    setOccLoadingMore(false);
    setOccError(null);
    setHasMore(false);
    setMineStatus({});
    setShowAllAnime(false);
    requestIdRef.current++;
  }, [wordId]);

  if (wordId === null) return null;

  const loadOccurrences = async (
    detail: NonNullable<typeof data>['detail'],
    offset: number,
    limit: number,
    append: boolean,
  ) => {
    const reqId = ++requestIdRef.current;
    if (append) {
      setOccLoadingMore(true);
    } else {
      setOccLoading(true);
      setOccError(null);
    }
    try {
      const rows = await apiClient.getWordOccurrences(
        detail.headword,
        detail.word,
        detail.reading,
        limit,
        offset,
      );
      if (reqId !== requestIdRef.current) return;
      setOccurrences((prev) => (append ? [...prev, ...rows] : rows));
      setHasMore(rows.length === limit);
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
    void loadOccurrences(data.detail, 0, INITIAL_PAGE_SIZE, false);
  };

  const handleLoadMore = () => {
    if (!data || occLoadingMore || !hasMore) return;
    void loadOccurrences(data.detail, occurrences.length, LOAD_MORE_SIZE, true);
  };

  const handleMine = async (
    occ: VocabularyOccurrenceEntry,
    mode: 'word' | 'sentence' | 'audio',
  ) => {
    const params = buildStatsMineCardParams(occ, data!.detail.headword, mode);
    if (!params) {
      return;
    }

    const key = `${occ.sessionId}-${occ.lineIndex}-${occ.segmentStartMs}-${mode}`;
    setMineStatus((prev) => ({ ...prev, [key]: { loading: true } }));
    try {
      const result = await apiClient.mineCard(params);
      const responseError = getStatsMineCardError(result);
      if (responseError) {
        setMineStatus((prev) => ({ ...prev, [key]: { error: responseError } }));
      } else {
        setMineStatus((prev) => ({ ...prev, [key]: { success: true } }));
        const label =
          mode === 'audio'
            ? 'Audio card'
            : mode === 'word'
              ? data!.detail.headword
              : occ.text.slice(0, 30);
        if (typeof Notification !== 'undefined' && Notification.permission === 'granted') {
          new Notification('Anki Card Created', { body: `Mined: ${label}`, icon: '/favicon.png' });
        } else if (typeof Notification !== 'undefined' && Notification.permission !== 'denied') {
          Notification.requestPermission().then((p) => {
            if (p === 'granted') new Notification('Anki Card Created', { body: `Mined: ${label}` });
          });
        }
      }
    } catch (err) {
      setMineStatus((prev) => ({
        ...prev,
        [key]: { error: err instanceof Error ? err.message : String(err) },
      }));
    }
  };

  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center p-4">
      <button
        type="button"
        aria-label="Close word detail panel"
        className="absolute inset-0 bg-ctp-crust/70 backdrop-blur-[2px]"
        onClick={onClose}
      />
      <aside className="relative flex max-h-[85vh] w-full max-w-2xl flex-col overflow-hidden rounded-xl border border-ctp-surface1 bg-ctp-mantle shadow-2xl">
        <div className="flex min-h-0 flex-1 flex-col">
          <div className="flex items-start justify-between border-b border-ctp-surface1 px-5 py-4">
            <div className="min-w-0">
              <div className="text-xs uppercase tracking-[0.18em] text-ctp-overlay1">
                Word Detail
              </div>
              {loading && <div className="mt-2 text-sm text-ctp-overlay2">Loading...</div>}
              {error && <div className="mt-2 text-sm text-ctp-red">Error: {error}</div>}
              {data && (
                <>
                  <h2 className="mt-1 truncate text-3xl font-semibold text-ctp-text">
                    {data.detail.headword}
                  </h2>
                  <div className="mt-1 text-sm text-ctp-subtext0">
                    {fullReading(data.detail.headword, data.detail.reading) || data.detail.word}
                  </div>
                  <div className="mt-2 flex flex-wrap gap-1.5">
                    {data.detail.partOfSpeech && <PosBadge pos={data.detail.partOfSpeech} />}
                    {data.detail.pos1 && data.detail.pos1 !== data.detail.partOfSpeech && (
                      <span className="rounded-full bg-ctp-surface1 px-2 py-0.5 text-[11px] text-ctp-subtext0">
                        {data.detail.pos1}
                      </span>
                    )}
                    {data.detail.pos2 && (
                      <span className="rounded-full bg-ctp-surface1 px-2 py-0.5 text-[11px] text-ctp-subtext0">
                        {data.detail.pos2}
                      </span>
                    )}
                    {data.detail.pos3 && (
                      <span className="rounded-full bg-ctp-surface1 px-2 py-0.5 text-[11px] text-ctp-subtext0">
                        {data.detail.pos3}
                      </span>
                    )}
                  </div>
                </>
              )}
            </div>
            <div className="flex items-center gap-2">
              {data && onToggleExclusion && (
                <button
                  type="button"
                  className={`rounded-md border px-3 py-1.5 text-xs font-medium transition ${
                    isExcluded?.(data.detail)
                      ? 'border-ctp-red/50 bg-ctp-red/10 text-ctp-red hover:bg-ctp-red/20'
                      : 'border-ctp-surface2 text-ctp-subtext0 hover:border-ctp-red hover:text-ctp-red'
                  }`}
                  onClick={() => onToggleExclusion(data.detail)}
                >
                  {isExcluded?.(data.detail) ? 'Excluded' : 'Exclude'}
                </button>
              )}
              <button
                type="button"
                className="rounded-md border border-ctp-surface2 px-3 py-1.5 text-xs font-medium text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue"
                onClick={onClose}
              >
                Close
              </button>
            </div>
          </div>

          <div className="flex-1 overflow-y-auto px-5 py-4 space-y-5">
            {data && (
              <>
                <div className="grid grid-cols-3 gap-3">
                  <div className="rounded-lg bg-ctp-surface0 p-3 text-center">
                    <div className="text-lg font-bold text-ctp-blue">
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
                      Media Appearances
                    </h3>
                    <div className="space-y-1.5">
                      {(showAllAnime
                        ? data.animeAppearances
                        : data.animeAppearances.slice(0, MEDIA_APPEARANCES_LIMIT)
                      ).map((a) => (
                        <button
                          key={a.animeId}
                          type="button"
                          onClick={() => {
                            onClose();
                            onNavigateToAnime?.(a.animeId);
                          }}
                          className="w-full flex items-center justify-between rounded-lg bg-ctp-surface0 px-3 py-2 text-sm transition hover:border-ctp-blue hover:ring-1 hover:ring-ctp-blue text-left"
                        >
                          <span className="truncate text-ctp-text">{a.animeTitle}</span>
                          <span className="ml-2 shrink-0 rounded-full bg-ctp-blue/10 px-2 py-0.5 text-[11px] font-medium text-ctp-blue">
                            {formatNumber(a.occurrenceCount)}
                          </span>
                        </button>
                      ))}
                    </div>
                    {data.animeAppearances.length > MEDIA_APPEARANCES_LIMIT && (
                      <button
                        type="button"
                        className="mt-2 w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-blue hover:text-ctp-blue"
                        onClick={() => setShowAllAnime((prev) => !prev)}
                      >
                        {showAllAnime
                          ? 'Show less'
                          : `Show ${formatNumber(
                              data.animeAppearances.length - MEDIA_APPEARANCES_LIMIT,
                            )} more`}
                      </button>
                    )}
                  </section>
                )}

                {data.similarWords.length > 0 && (
                  <section>
                    <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">
                      Related Seen Words
                    </h3>
                    <div className="flex flex-wrap gap-1.5">
                      {data.similarWords.map((sw) => (
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
                  <h3 className="text-xs font-semibold uppercase tracking-wide text-ctp-overlay1 mb-2">
                    Example Lines
                  </h3>
                  {!occLoaded && !occLoading && (
                    <button
                      type="button"
                      className="w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-blue hover:text-ctp-blue"
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
                    <div className="text-sm text-ctp-overlay2">
                      No example lines tracked yet. Lines are stored for sessions recorded after the
                      subtitle tracking update.
                    </div>
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
                          <div className="mt-3 flex items-center gap-2 text-xs text-ctp-overlay1">
                            <span>
                              {formatSegment(occ.segmentStartMs)}-{formatSegment(occ.segmentEndMs)}{' '}
                              · session {occ.sessionId}
                            </span>
                            {(() => {
                              const unavailableReason = getStatsMineCardUnavailableReason(occ);
                              const baseKey = `${occ.sessionId}-${occ.lineIndex}-${occ.segmentStartMs}`;
                              const wordStatus = mineStatus[`${baseKey}-word`];
                              const sentenceStatus = mineStatus[`${baseKey}-sentence`];
                              const audioStatus = mineStatus[`${baseKey}-audio`];
                              return (
                                <>
                                  <button
                                    type="button"
                                    title={unavailableReason ?? 'Mine this word from video clip'}
                                    className="rounded border border-ctp-surface2 px-1.5 py-0.5 text-[10px] font-medium text-ctp-subtext0 transition hover:border-ctp-mauve hover:text-ctp-mauve disabled:cursor-not-allowed disabled:opacity-60"
                                    disabled={wordStatus?.loading || !!unavailableReason}
                                    onClick={() => void handleMine(occ, 'word')}
                                  >
                                    {wordStatus?.loading
                                      ? 'Mining...'
                                      : wordStatus?.success
                                        ? 'Mined!'
                                        : unavailableReason
                                          ? 'Unavailable'
                                          : 'Mine Word'}
                                  </button>
                                  <button
                                    type="button"
                                    title={
                                      unavailableReason ?? 'Mine this sentence from video clip'
                                    }
                                    className="rounded border border-ctp-surface2 px-1.5 py-0.5 text-[10px] font-medium text-ctp-subtext0 transition hover:border-ctp-green hover:text-ctp-green disabled:cursor-not-allowed disabled:opacity-60"
                                    disabled={sentenceStatus?.loading || !!unavailableReason}
                                    onClick={() => void handleMine(occ, 'sentence')}
                                  >
                                    {sentenceStatus?.loading
                                      ? 'Mining...'
                                      : sentenceStatus?.success
                                        ? 'Mined!'
                                        : unavailableReason
                                          ? 'Unavailable'
                                          : 'Mine Sentence'}
                                  </button>
                                  <button
                                    type="button"
                                    title={unavailableReason ?? 'Mine this line as audio-only card'}
                                    className="rounded border border-ctp-surface2 px-1.5 py-0.5 text-[10px] font-medium text-ctp-subtext0 transition hover:border-ctp-blue hover:text-ctp-blue disabled:cursor-not-allowed disabled:opacity-60"
                                    disabled={audioStatus?.loading || !!unavailableReason}
                                    onClick={() => void handleMine(occ, 'audio')}
                                  >
                                    {audioStatus?.loading
                                      ? 'Mining...'
                                      : audioStatus?.success
                                        ? 'Mined!'
                                        : unavailableReason
                                          ? 'Unavailable'
                                          : 'Mine Audio'}
                                  </button>
                                </>
                              );
                            })()}
                          </div>
                          {(() => {
                            const baseKey = `${occ.sessionId}-${occ.lineIndex}-${occ.segmentStartMs}`;
                            const errors = ['word', 'sentence', 'audio']
                              .map((m) => mineStatus[`${baseKey}-${m}`]?.error)
                              .filter(Boolean);
                            return errors.length > 0 ? (
                              <div className="mt-1 text-[10px] text-ctp-red">{errors[0]}</div>
                            ) : null;
                          })()}
                          <p className="mt-3 rounded-lg bg-ctp-base/70 px-3 py-3 text-sm leading-6 text-ctp-text">
                            {highlightWord(occ.text, [data!.detail.headword, data!.detail.word])}
                          </p>
                        </article>
                      ))}
                      {hasMore && (
                        <button
                          type="button"
                          className="w-full rounded-lg border border-ctp-surface2 bg-ctp-surface0 px-4 py-2 text-sm font-medium text-ctp-text transition hover:border-ctp-blue hover:text-ctp-blue disabled:cursor-not-allowed disabled:opacity-60"
                          onClick={handleLoadMore}
                          disabled={occLoadingMore}
                        >
                          {occLoadingMore ? 'Loading more...' : 'Load more'}
                        </button>
                      )}
                    </div>
                  )}
                </section>
              </>
            )}
          </div>
        </div>
      </aside>
    </div>
  );
}
