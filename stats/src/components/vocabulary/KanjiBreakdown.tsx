import { useMemo } from 'react';
import type { KanjiEntry } from '../../types/stats';
import { formatNumber } from '../../lib/formatters';

interface KanjiBreakdownProps {
  kanji: KanjiEntry[];
  selectedKanjiId?: number | null;
  onSelectKanji?: (entry: KanjiEntry) => void;
}

// Heat scale from rare (cool) to very frequent (warm). Catppuccin Macchiato.
const FREQ_TIERS = [
  { min: 0.85, color: 'text-ctp-peach', swatch: 'bg-ctp-peach', label: 'Very frequent' },
  { min: 0.6, color: 'text-ctp-yellow', swatch: 'bg-ctp-yellow', label: 'Frequent' },
  { min: 0.35, color: 'text-ctp-green', swatch: 'bg-ctp-green', label: 'Common' },
  { min: 0.15, color: 'text-ctp-teal', swatch: 'bg-ctp-teal', label: 'Occasional' },
  { min: 0, color: 'text-ctp-subtext0', swatch: 'bg-ctp-subtext0', label: 'Rare' },
] as const;

function tierFor(intensity: number) {
  return FREQ_TIERS.find((tier) => intensity >= tier.min) ?? FREQ_TIERS[FREQ_TIERS.length - 1]!;
}

export function KanjiBreakdown({
  kanji,
  selectedKanjiId = null,
  onSelectKanji,
}: KanjiBreakdownProps) {
  const { totalOccurrences, maxLogFreq } = useMemo(() => {
    let total = 0;
    let maxFreq = 1;
    for (const entry of kanji) {
      total += entry.frequency;
      if (entry.frequency > maxFreq) maxFreq = entry.frequency;
    }
    return { totalOccurrences: total, maxLogFreq: Math.log(maxFreq + 1) };
  }, [kanji]);

  if (kanji.length === 0) return null;

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <h3 className="text-sm font-semibold text-ctp-text">
          Kanji Encountered
          <span className="ml-2 font-normal text-ctp-subtext0">
            {formatNumber(kanji.length)} unique · {formatNumber(totalOccurrences)} seen
          </span>
        </h3>
        <div className="flex items-center gap-1.5 text-[11px] text-ctp-subtext0">
          <span>rare</span>
          <div className="flex items-center gap-1">
            {[...FREQ_TIERS].reverse().map((tier) => (
              <span
                key={tier.label}
                className={`h-2 w-2 rounded-full ${tier.swatch}`}
                title={tier.label}
              />
            ))}
          </div>
          <span>frequent</span>
        </div>
      </div>
      <div className="flex flex-wrap gap-1">
        {kanji.map((k) => {
          // Log scale keeps the heavily-skewed frequency distribution readable.
          const intensity = maxLogFreq > 0 ? Math.log(k.frequency + 1) / maxLogFreq : 0;
          const tier = tierFor(intensity);
          const selected = selectedKanjiId === k.kanjiId;
          return (
            <button
              type="button"
              key={k.kanji}
              className={`cursor-pointer rounded-md px-1.5 py-0.5 text-xl leading-none font-medium transition-colors duration-150 ${tier.color} ${
                selected ? 'bg-ctp-surface2 ring-1 ring-ctp-lavender' : 'hover:bg-ctp-surface1'
              }`}
              title={`${k.kanji} — seen ${formatNumber(k.frequency)}×`}
              aria-label={`${k.kanji} — seen ${k.frequency} times`}
              aria-pressed={selected}
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
