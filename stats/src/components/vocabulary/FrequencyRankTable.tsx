import { useMemo, useState } from 'react';
import { PosBadge } from './pos-helpers';
import { fullReading } from '../../lib/reading-utils';
import type { VocabularyEntry } from '../../types/stats';

interface FrequencyRankTableProps {
  words: VocabularyEntry[];
  knownWords: Set<string>;
  onSelectWord?: (word: VocabularyEntry) => void;
}

const PAGE_SIZE = 25;

export function FrequencyRankTable({ words, knownWords, onSelectWord }: FrequencyRankTableProps) {
  const [page, setPage] = useState(0);
  const [hideKnown, setHideKnown] = useState(true);
  const [collapsed, setCollapsed] = useState(false);

  const hasKnownData = knownWords.size > 0;

  const isWordKnown = (w: VocabularyEntry): boolean => {
    return knownWords.has(w.headword) || knownWords.has(w.word);
  };

  const ranked = useMemo(() => {
    let filtered = words.filter((w) => w.frequencyRank != null && w.frequencyRank > 0);
    if (hideKnown && hasKnownData) {
      filtered = filtered.filter((w) => !isWordKnown(w));
    }

    const byHeadword = new Map<string, VocabularyEntry>();
    for (const w of filtered) {
      const existing = byHeadword.get(w.headword);
      if (!existing) {
        byHeadword.set(w.headword, { ...w });
      } else {
        existing.frequency += w.frequency;
        existing.animeCount = Math.max(existing.animeCount, w.animeCount);
        if (w.frequencyRank! < existing.frequencyRank!) {
          existing.frequencyRank = w.frequencyRank;
        }
        if (!existing.reading && w.reading) {
          existing.reading = w.reading;
        }
        if (!existing.partOfSpeech && w.partOfSpeech) {
          existing.partOfSpeech = w.partOfSpeech;
        }
      }
    }

    return [...byHeadword.values()].sort((a, b) => a.frequencyRank! - b.frequencyRank!);
  }, [words, knownWords, hideKnown, hasKnownData]);

  if (words.every((w) => w.frequencyRank == null)) {
    return (
      <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
        <h3 className="text-sm font-semibold text-ctp-text mb-2">Most Common Words Seen</h3>
        <div className="text-xs text-ctp-overlay2">
          No frequency rank data available. Run the frequency backfill script or install a frequency
          dictionary.
        </div>
      </div>
    );
  }

  const totalPages = Math.ceil(ranked.length / PAGE_SIZE);
  const paged = ranked.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <div className="flex items-center justify-between">
        <button
          type="button"
          onClick={() => setCollapsed(!collapsed)}
          className="flex items-center gap-2 text-sm font-semibold text-ctp-text hover:text-ctp-subtext1 transition-colors"
        >
          <span
            className={`text-xs text-ctp-overlay2 transition-transform ${collapsed ? '' : 'rotate-90'}`}
          >
            {'\u25B6'}
          </span>
          {hideKnown && hasKnownData ? 'Common Words Not Yet Mined' : 'Most Common Words Seen'}
        </button>
        <div className="flex items-center gap-3">
          {hasKnownData && (
            <button
              type="button"
              onClick={() => {
                setHideKnown(!hideKnown);
                setPage(0);
              }}
              className={`px-2.5 py-1 rounded-lg text-xs transition-colors border ${
                hideKnown
                  ? 'bg-ctp-surface2 text-ctp-text border-ctp-blue/50'
                  : 'bg-ctp-surface0 text-ctp-overlay2 border-ctp-surface1 hover:text-ctp-subtext0'
              }`}
            >
              Hide Known
            </button>
          )}
          <span className="text-xs text-ctp-overlay2">{ranked.length} words</span>
        </div>
      </div>
      {collapsed ? null : ranked.length === 0 ? (
        <div className="text-xs text-ctp-overlay2 mt-3">
          {hideKnown ? 'All ranked words are already in Anki!' : 'No words with frequency data.'}
        </div>
      ) : (
        <>
          <div className="overflow-x-auto mt-3">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-xs text-ctp-overlay2 border-b border-ctp-surface1">
                  <th className="text-left py-2 pr-3 font-medium w-16">Rank</th>
                  <th className="text-left py-2 pr-3 font-medium">Word</th>
                  <th className="text-left py-2 pr-3 font-medium w-20">POS</th>
                  <th className="text-right py-2 font-medium w-20">Seen</th>
                </tr>
              </thead>
              <tbody>
                {paged.map((w) => (
                  <tr
                    key={w.wordId}
                    onClick={() => onSelectWord?.(w)}
                    className="border-b border-ctp-surface1 last:border-0 cursor-pointer hover:bg-ctp-surface1/50 transition-colors"
                  >
                    <td className="py-1.5 pr-3 font-mono tabular-nums text-ctp-peach text-xs">
                      #{w.frequencyRank!.toLocaleString()}
                    </td>
                    <td className="py-1.5 pr-3">
                      <span className="text-ctp-text font-medium">{w.headword}</span>
                      {(() => {
                        const reading = fullReading(w.headword, w.reading);
                        if (!reading || reading === w.headword) return null;
                        return (
                          <span className="text-ctp-subtext0 text-xs ml-1.5">
                            【{reading}】
                          </span>
                        );
                      })()}
                    </td>
                    <td className="py-1.5 pr-3">
                      {w.partOfSpeech && <PosBadge pos={w.partOfSpeech} />}
                    </td>
                    <td className="py-1.5 text-right font-mono tabular-nums text-ctp-blue text-xs">
                      {w.frequency}x
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {totalPages > 1 && (
            <div className="flex items-center justify-center gap-3 mt-3 text-xs">
              <button
                type="button"
                disabled={page === 0}
                onClick={() => setPage(page - 1)}
                className="px-2 py-1 rounded bg-ctp-surface1 text-ctp-text disabled:opacity-30 hover:bg-ctp-surface2 transition-colors"
              >
                Prev
              </button>
              <span className="text-ctp-overlay2">
                {page + 1} / {totalPages}
              </span>
              <button
                type="button"
                disabled={page >= totalPages - 1}
                onClick={() => setPage(page + 1)}
                className="px-2 py-1 rounded bg-ctp-surface1 text-ctp-text disabled:opacity-30 hover:bg-ctp-surface2 transition-colors"
              >
                Next
              </button>
            </div>
          )}
        </>
      )}
    </div>
  );
}
