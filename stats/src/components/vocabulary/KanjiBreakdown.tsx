import type { KanjiEntry } from '../../types/stats';

interface KanjiBreakdownProps {
  kanji: KanjiEntry[];
  selectedKanjiId?: number | null;
  onSelectKanji?: (entry: KanjiEntry) => void;
}

export function KanjiBreakdown({
  kanji,
  selectedKanjiId = null,
  onSelectKanji,
}: KanjiBreakdownProps) {
  if (kanji.length === 0) return null;

  const maxFreq = kanji.reduce((max, entry) => Math.max(max, entry.frequency), 1);

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-sm font-semibold text-ctp-text mb-3">Kanji Encountered</h3>
      <div className="flex flex-wrap gap-1">
        {kanji.map((k) => {
          const ratio = k.frequency / maxFreq;
          const opacity = Math.max(0.3, ratio);
          return (
            <button
              type="button"
              key={k.kanji}
              className={`cursor-pointer rounded px-1 text-lg text-ctp-teal transition ${
                selectedKanjiId === k.kanjiId
                  ? 'bg-ctp-teal/10 ring-1 ring-ctp-teal'
                  : 'hover:bg-ctp-surface1/80'
              }`}
              style={{ opacity }}
              title={`${k.kanji} — seen ${k.frequency}x`}
              aria-label={`${k.kanji} — seen ${k.frequency} times`}
              onClick={() => onSelectKanji?.(k)}
            >
              {k.kanji}
            </button>
          );
        })}
      </div>
    </div>
  );
}
