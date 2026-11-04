import { useMemo, useState } from 'react';
import { formatNumber } from '../../lib/formatters';
import type { KanjiEntry } from '../../types/stats';

interface KanjiBreakdownProps {
  kanji: KanjiEntry[];
  knownWords: ReadonlySet<string>;
  selectedKanjiId?: number | null;
  onSelectKanji?: (entry: KanjiEntry) => void;
}

type KanjiFilter = 'all' | 'unknown';

// Rendering every kanji at once is slow to lay out and rarely what you want first.
const INITIAL_VISIBLE = 180;
const MIN_CELL_ALPHA = 0.08;
const MAX_CELL_ALPHA = 0.6;
const KANJI_CHARACTER = /\p{Script=Han}/u;

/** Every kanji appearing in at least one known word. */
export function kanjiInKnownWords(knownWords: Iterable<string>): Set<string> {
  const found = new Set<string>();
  for (const word of knownWords) {
    for (const character of word) {
      if (KANJI_CHARACTER.test(character)) found.add(character);
    }
  }
  return found;
}

/** Square root, so common kanji still differ while rare ones stay visible. */
export function kanjiCellAlpha(frequency: number, maxFrequency: number): number {
  const intensity = maxFrequency > 0 ? Math.sqrt(frequency / maxFrequency) : 0;
  return MIN_CELL_ALPHA + (MAX_CELL_ALPHA - MIN_CELL_ALPHA) * intensity;
}

export function KanjiBreakdown({
  kanji,
  knownWords,
  selectedKanjiId = null,
  onSelectKanji,
}: KanjiBreakdownProps) {
  const [filter, setFilter] = useState<KanjiFilter>('all');
  const [showAll, setShowAll] = useState(false);
  const hasKnownData = knownWords.size > 0;
  const knownKanji = useMemo(() => kanjiInKnownWords(knownWords), [knownWords]);
  const unknownKanji = useMemo(
    () => kanji.filter((entry) => !knownKanji.has(entry.kanji)),
    [kanji, knownKanji],
  );
  const maxFrequency = useMemo(
    () => kanji.reduce((max, entry) => Math.max(max, entry.frequency), 1),
    [kanji],
  );

  if (kanji.length === 0) return null;

  const matching = filter === 'unknown' ? unknownKanji : kanji;
  const visible = showAll ? matching : matching.slice(0, INITIAL_VISIBLE);

  return (
    <section
      aria-label="Kanji"
      className="rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4"
    >
      <div className="mb-3 flex flex-wrap items-start justify-between gap-2">
        <div>
          <h3 className="text-xs font-semibold text-ctp-text">Kanji</h3>
          <p className="mt-0.5 text-[11px] text-ctp-overlay2">
            Shaded by how often you have seen it.
            {hasKnownData ? ' Dashed means it is not in any word you know.' : ''}
          </p>
        </div>
        {hasKnownData ? (
          <div
            role="group"
            aria-label="Kanji filter"
            className="flex rounded-lg border border-ctp-surface1 p-0.5 text-[11px]"
          >
            {(
              [
                ['all', `All ${formatNumber(kanji.length)}`],
                ['unknown', `Not in a known word ${formatNumber(unknownKanji.length)}`],
              ] as const
            ).map(([value, label]) => (
              <button
                key={value}
                type="button"
                aria-pressed={filter === value}
                onClick={() => setFilter(value)}
                className={`rounded-md px-2 py-0.5 transition-colors ${
                  filter === value
                    ? 'bg-ctp-surface2 text-ctp-text'
                    : 'text-ctp-overlay2 hover:text-ctp-subtext0'
                }`}
              >
                {label}
              </button>
            ))}
          </div>
        ) : null}
      </div>
      <div className="grid grid-cols-[repeat(auto-fill,minmax(2.25rem,1fr))] gap-1">
        {visible.map((entry) => {
          const selected = selectedKanjiId === entry.kanjiId;
          const outsideKnown = hasKnownData && !knownKanji.has(entry.kanji);
          return (
            <button
              type="button"
              key={entry.kanji}
              lang="ja"
              title={`${entry.kanji}: seen ${formatNumber(entry.frequency)} times`}
              aria-label={`${entry.kanji}, seen ${entry.frequency} times`}
              aria-pressed={selected}
              onClick={() => onSelectKanji?.(entry)}
              style={{
                backgroundColor: `rgb(138 173 244 / ${kanjiCellAlpha(entry.frequency, maxFrequency).toFixed(2)})`,
              }}
              className={`flex aspect-square items-center justify-center rounded text-lg leading-none text-ctp-text transition hover:ring-1 hover:ring-ctp-blue ${
                selected ? 'ring-1 ring-ctp-lavender' : ''
              } ${outsideKnown ? 'outline-1 -outline-offset-1 outline-ctp-blue/70 outline-dashed' : ''}`}
            >
              {entry.kanji}
            </button>
          );
        })}
      </div>
      {matching.length > INITIAL_VISIBLE ? (
        <button
          type="button"
          onClick={() => setShowAll((value) => !value)}
          className="mt-2 text-[11px] text-ctp-overlay2 transition-colors hover:text-ctp-blue"
        >
          {showAll ? 'Show fewer' : `Show all ${formatNumber(matching.length)}`}
        </button>
      ) : null}
    </section>
  );
}
