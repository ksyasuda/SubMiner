import { useEffect, useRef, useState } from 'react';
import { apiClient } from '../../lib/api-client';
import {
  getSentenceSearchMineAvailability,
  renderSentenceWithMatches,
} from '../../lib/sentence-search';
import type { SentenceSearchResult } from '../../types/stats';

const SEARCH_LIMIT = 50;
const SEARCH_DEBOUNCE_MS = 160;

type MineMode = 'word' | 'sentence' | 'audio';
type MineStatus = { loading?: boolean; success?: boolean; error?: string };

function formatSegment(ms: number | null): string {
  if (ms == null || !Number.isFinite(ms)) return '--:--';
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${String(seconds).padStart(2, '0')}`;
}

function resultKey(result: SentenceSearchResult, index: number): string {
  return `${result.sessionId}-${result.lineIndex}-${result.segmentStartMs ?? index}`;
}

function statusKey(result: SentenceSearchResult, index: number, mode: MineMode): string {
  return `${resultKey(result, index)}-${mode}`;
}

function buttonLabel(
  mode: MineMode,
  status: MineStatus | undefined,
  disabledLabel: string,
): string {
  if (status?.loading) return 'Mining...';
  if (status?.success) return 'Mined!';
  if (disabledLabel) return disabledLabel;
  if (mode === 'word') return 'Word';
  if (mode === 'audio') return 'Audio';
  return 'Sentence';
}

export function SearchTab() {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<SentenceSearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [mineStatus, setMineStatus] = useState<Record<string, MineStatus>>({});
  const requestRef = useRef(0);

  useEffect(() => {
    const trimmed = query.trim();
    const requestId = ++requestRef.current;
    setMineStatus({});

    if (!trimmed) {
      setResults([]);
      setLoading(false);
      setError(null);
      return;
    }

    setLoading(true);
    setError(null);
    const timer = window.setTimeout(() => {
      apiClient
        .searchSentences(trimmed, SEARCH_LIMIT)
        .then((nextResults) => {
          if (requestId !== requestRef.current) return;
          setResults(nextResults);
        })
        .catch((err: Error) => {
          if (requestId !== requestRef.current) return;
          setError(err.message);
          setResults([]);
        })
        .finally(() => {
          if (requestId !== requestRef.current) return;
          setLoading(false);
        });
    }, SEARCH_DEBOUNCE_MS);

    return () => {
      window.clearTimeout(timer);
    };
  }, [query]);

  const handleMine = async (
    result: SentenceSearchResult,
    index: number,
    mode: MineMode,
  ): Promise<void> => {
    const availability = getSentenceSearchMineAvailability(result, query);
    if (mode === 'sentence' ? !availability.canMineSentence : !availability.canMineWordAudio) {
      return;
    }
    if (!result.sourcePath || result.segmentStartMs == null || result.segmentEndMs == null) {
      return;
    }

    const key = statusKey(result, index, mode);
    setMineStatus((prev) => ({ ...prev, [key]: { loading: true } }));
    try {
      const searchedWord = availability.exactMatch ? query.trim() : '';
      const response = await apiClient.mineCard({
        sourcePath: result.sourcePath,
        startMs: result.segmentStartMs,
        endMs: result.segmentEndMs,
        sentence: result.text,
        word: searchedWord,
        secondaryText: result.secondaryText,
        videoTitle: result.videoTitle,
        mode,
      });

      if (response.error) {
        setMineStatus((prev) => ({ ...prev, [key]: { error: response.error } }));
        return;
      }
      setMineStatus((prev) => ({ ...prev, [key]: { success: true } }));
    } catch (err) {
      setMineStatus((prev) => ({
        ...prev,
        [key]: { error: err instanceof Error ? err.message : String(err) },
      }));
    }
  };

  const trimmedQuery = query.trim();

  return (
    <div className="space-y-4">
      <section className="rounded-lg border border-ctp-surface1 bg-ctp-mantle/70 p-4">
        <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
          <label className="min-w-0 flex-1">
            <span className="mb-2 block text-xs font-semibold uppercase tracking-[0.18em] text-ctp-overlay1">
              Sentence Search
            </span>
            <input
              type="search"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Search sentence text or media..."
              className="w-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 px-4 py-3 text-base text-ctp-text placeholder:text-ctp-overlay2 focus:border-ctp-yellow focus:outline-none"
              autoComplete="off"
            />
          </label>
          <div className="rounded-lg border border-ctp-surface1 bg-ctp-surface0 px-4 py-2 text-right">
            <div className="text-xl font-semibold text-ctp-yellow">
              {loading ? '...' : results.length}
            </div>
            <div className="text-[11px] uppercase tracking-wide text-ctp-overlay1">Matches</div>
          </div>
        </div>
      </section>

      {error && (
        <div className="rounded-lg border border-ctp-red/40 bg-ctp-red/10 p-3 text-sm text-ctp-red">
          Error: {error}
        </div>
      )}

      {!trimmedQuery && (
        <div className="rounded-lg border border-ctp-surface1 bg-ctp-surface0/70 p-6 text-sm text-ctp-overlay2">
          Search your tracked subtitle lines.
        </div>
      )}

      {trimmedQuery && !loading && !error && results.length === 0 && (
        <div className="rounded-lg border border-ctp-surface1 bg-ctp-surface0/70 p-6 text-sm text-ctp-overlay2">
          No sentence matches.
        </div>
      )}

      {results.length > 0 && (
        <div className="grid gap-3">
          {results.map((result, index) => {
            const availability = getSentenceSearchMineAvailability(result, trimmedQuery);
            const wordStatus = mineStatus[statusKey(result, index, 'word')];
            const sentenceStatus = mineStatus[statusKey(result, index, 'sentence')];
            const audioStatus = mineStatus[statusKey(result, index, 'audio')];
            const wordAudioDisabledReason =
              availability.unavailableReason ??
              (availability.exactMatch ? '' : 'Exact searched word not found in sentence.');
            const sentenceDisabledReason = availability.unavailableReason ?? '';
            const errors = [wordStatus?.error, sentenceStatus?.error, audioStatus?.error].filter(
              Boolean,
            );

            return (
              <article
                key={resultKey(result, index)}
                className="rounded-lg border border-ctp-surface1 bg-ctp-surface0/90 p-4 shadow-lg shadow-ctp-crust/20"
              >
                <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
                  <div className="min-w-0">
                    <div className="truncate text-sm font-semibold text-ctp-text">
                      {result.animeTitle ?? result.videoTitle}
                    </div>
                    <div className="mt-1 flex flex-wrap gap-x-2 gap-y-1 text-xs text-ctp-overlay1">
                      <span className="max-w-full truncate">{result.videoTitle}</span>
                      <span>line {result.lineIndex}</span>
                      <span>session {result.sessionId}</span>
                      <span>
                        {formatSegment(result.segmentStartMs)}-{formatSegment(result.segmentEndMs)}
                      </span>
                    </div>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {availability.exactMatch && (
                      <button
                        type="button"
                        title={wordAudioDisabledReason || 'Create a word card from this sentence'}
                        className="rounded-md border border-ctp-mauve/50 px-3 py-1.5 text-xs font-medium text-ctp-mauve transition hover:bg-ctp-mauve/10 disabled:cursor-not-allowed disabled:border-ctp-surface2 disabled:text-ctp-overlay1 disabled:opacity-60"
                        disabled={wordStatus?.loading || !availability.canMineWordAudio}
                        onClick={() => void handleMine(result, index, 'word')}
                      >
                        {buttonLabel(
                          'word',
                          wordStatus,
                          availability.canMineWordAudio ? '' : 'Unavailable',
                        )}
                      </button>
                    )}
                    <button
                      type="button"
                      title={sentenceDisabledReason || 'Create a sentence card from this line'}
                      className="rounded-md border border-ctp-green/50 px-3 py-1.5 text-xs font-medium text-ctp-green transition hover:bg-ctp-green/10 disabled:cursor-not-allowed disabled:border-ctp-surface2 disabled:text-ctp-overlay1 disabled:opacity-60"
                      disabled={sentenceStatus?.loading || !availability.canMineSentence}
                      onClick={() => void handleMine(result, index, 'sentence')}
                    >
                      {buttonLabel(
                        'sentence',
                        sentenceStatus,
                        availability.canMineSentence ? '' : 'Unavailable',
                      )}
                    </button>
                    {availability.exactMatch && (
                      <button
                        type="button"
                        title={wordAudioDisabledReason || 'Create an audio card from this sentence'}
                        className="rounded-md border border-ctp-blue/50 px-3 py-1.5 text-xs font-medium text-ctp-blue transition hover:bg-ctp-blue/10 disabled:cursor-not-allowed disabled:border-ctp-surface2 disabled:text-ctp-overlay1 disabled:opacity-60"
                        disabled={audioStatus?.loading || !availability.canMineWordAudio}
                        onClick={() => void handleMine(result, index, 'audio')}
                      >
                        {buttonLabel(
                          'audio',
                          audioStatus,
                          availability.canMineWordAudio ? '' : 'Unavailable',
                        )}
                      </button>
                    )}
                  </div>
                </div>

                <p className="mt-4 rounded-lg bg-ctp-base/70 px-4 py-3 text-base leading-7 text-ctp-text">
                  {renderSentenceWithMatches(result.text, trimmedQuery)}
                </p>
                {errors.length > 0 && <div className="mt-2 text-xs text-ctp-red">{errors[0]}</div>}
              </article>
            );
          })}
        </div>
      )}
    </div>
  );
}
