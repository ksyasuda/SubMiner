import { useMemo, useState } from 'react';
import { PosBadge } from './pos-helpers';
import { isKanaOnlyTokenText } from '../../lib/kana-token';
import { readBooleanPreference, writeBooleanPreference } from '../../lib/preference-storage';
import { fullReading } from '../../lib/reading-utils';
import type { VocabularyEntry } from '../../types/stats';

interface CrossAnimeWordsTableProps {
  words: VocabularyEntry[];
  knownWords: Set<string>;
  onSelectWord?: (word: VocabularyEntry) => void;
}

const PAGE_SIZE = 25;
const HIDE_KANA_ONLY_STORAGE_KEY = 'subminer.stats.crossAnimeWords.hideKanaOnly';

interface CrossAnimeWordsOptions {
  hideKnown: boolean;
  hideKanaOnly: boolean;
}

function isWordKnown(w: VocabularyEntry, knownWords: Set<string>): boolean {
  return knownWords.has(w.headword) || knownWords.has(w.word);
}

function isKanaOnlyWord(w: VocabularyEntry): boolean {
  return isKanaOnlyTokenText(w.headword || w.word);
}

export function buildCrossAnimeWordRows(
  words: VocabularyEntry[],
  knownWords: Set<string>,
  options: CrossAnimeWordsOptions,
): VocabularyEntry[] {
  const hasKnownData = knownWords.size > 0;
  let filtered = words.filter((w) => w.animeCount >= 2);
  if (options.hideKnown && hasKnownData) {
    filtered = filtered.filter((w) => !isWordKnown(w, knownWords));
  }
  if (options.hideKanaOnly) {
    filtered = filtered.filter((w) => !isKanaOnlyWord(w));
  }

  const byHeadword = new Map<string, VocabularyEntry>();
  for (const w of filtered) {
    const existing = byHeadword.get(w.headword);
    if (!existing) {
      byHeadword.set(w.headword, { ...w });
    } else {
      existing.frequency += w.frequency;
      existing.animeCount = Math.max(existing.animeCount, w.animeCount);
      if (
        w.frequencyRank != null &&
        (existing.frequencyRank == null || w.frequencyRank < existing.frequencyRank)
      ) {
        existing.frequencyRank = w.frequencyRank;
      }
      if (!existing.reading && w.reading) existing.reading = w.reading;
      if (!existing.partOfSpeech && w.partOfSpeech) existing.partOfSpeech = w.partOfSpeech;
    }
  }

  return [...byHeadword.values()].sort((a, b) => {
    if (b.animeCount !== a.animeCount) return b.animeCount - a.animeCount;
    return b.frequency - a.frequency;
  });
}

export function CrossAnimeWordsTable({
  words,
  knownWords,
  onSelectWord,
}: CrossAnimeWordsTableProps) {
  const [page, setPage] = useState(0);
  const [hideKnown, setHideKnown] = useState(true);
  const [hideKanaOnly, setHideKanaOnly] = useState(() =>
    readBooleanPreference(HIDE_KANA_ONLY_STORAGE_KEY, false),
  );
  const [collapsed, setCollapsed] = useState(false);

  const hasKnownData = knownWords.size > 0;

  const ranked = useMemo(() => {
    return buildCrossAnimeWordRows(words, knownWords, { hideKnown, hideKanaOnly });
  }, [words, knownWords, hideKnown, hideKanaOnly]);

  const hasMultiAnimeWords = words.some((w) => w.animeCount >= 2);
  if (!hasMultiAnimeWords) return null;

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
          Words Across Multiple Titles
        </button>
        <div className="flex flex-wrap items-center justify-end gap-2">
          {hasKnownData && (
            <button
              type="button"
              aria-pressed={hideKnown}
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
          <button
            type="button"
            aria-pressed={hideKanaOnly}
            onClick={() => {
              const next = !hideKanaOnly;
              setHideKanaOnly(next);
              writeBooleanPreference(HIDE_KANA_ONLY_STORAGE_KEY, next);
              setPage(0);
            }}
            className={`px-2.5 py-1 rounded-lg text-xs transition-colors border ${
              hideKanaOnly
                ? 'bg-ctp-surface2 text-ctp-text border-ctp-blue/50'
                : 'bg-ctp-surface0 text-ctp-overlay2 border-ctp-surface1 hover:text-ctp-subtext0'
            }`}
          >
            Hide Kana
          </button>
          <span className="text-xs text-ctp-overlay2">{ranked.length} words</span>
        </div>
      </div>
      {collapsed ? null : ranked.length === 0 ? (
        <div className="text-xs text-ctp-overlay2 mt-3">
          {hideKnown && hasKnownData && !hideKanaOnly
            ? 'All words that span multiple titles are already known!'
            : (hideKnown && hasKnownData) || hideKanaOnly
              ? 'No words across multiple titles match the active filters.'
              : 'No words found across multiple titles.'}
        </div>
      ) : (
        <>
          <div className="overflow-x-auto mt-3">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-xs text-ctp-overlay2 border-b border-ctp-surface1">
                  <th className="text-left py-2 pr-3 font-medium">Word</th>
                  <th className="text-left py-2 pr-3 font-medium">Reading</th>
                  <th className="text-left py-2 pr-3 font-medium w-20">POS</th>
                  <th className="text-right py-2 pr-3 font-medium w-16">Titles</th>
                  <th className="text-right py-2 font-medium w-16">Seen</th>
                </tr>
              </thead>
              <tbody>
                {paged.map((w) => (
                  <tr
                    key={w.wordId}
                    onClick={() => onSelectWord?.(w)}
                    className="border-b border-ctp-surface1 last:border-0 cursor-pointer hover:bg-ctp-surface1/50 transition-colors"
                  >
                    <td className="py-1.5 pr-3 text-ctp-text font-medium">{w.headword}</td>
                    <td className="py-1.5 pr-3 text-ctp-subtext0">
                      {fullReading(w.headword, w.reading) || w.headword}
                    </td>
                    <td className="py-1.5 pr-3">
                      {w.partOfSpeech && <PosBadge pos={w.partOfSpeech} />}
                    </td>
                    <td className="py-1.5 pr-3 text-right font-mono tabular-nums text-ctp-green text-xs">
                      {w.animeCount}
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
