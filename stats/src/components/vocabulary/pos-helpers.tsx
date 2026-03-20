import type { VocabularyEntry } from '../../types/stats';

const POS_COLORS: Record<string, string> = {
  noun: 'bg-ctp-blue/15 text-ctp-blue',
  verb: 'bg-ctp-green/15 text-ctp-green',
  adjective: 'bg-ctp-mauve/15 text-ctp-mauve',
  adverb: 'bg-ctp-peach/15 text-ctp-peach',
  particle: 'bg-ctp-overlay0/15 text-ctp-overlay0',
  auxiliary_verb: 'bg-ctp-overlay0/15 text-ctp-overlay0',
  conjunction: 'bg-ctp-overlay0/15 text-ctp-overlay0',
  prenominal: 'bg-ctp-yellow/15 text-ctp-yellow',
  suffix: 'bg-ctp-flamingo/15 text-ctp-flamingo',
  prefix: 'bg-ctp-flamingo/15 text-ctp-flamingo',
  interjection: 'bg-ctp-rosewater/15 text-ctp-rosewater',
};

const DEFAULT_POS_COLOR = 'bg-ctp-surface1 text-ctp-subtext0';

export function posColor(pos: string): string {
  return POS_COLORS[pos] ?? DEFAULT_POS_COLOR;
}

export function PosBadge({ pos }: { pos: string }) {
  return (
    <span className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${posColor(pos)}`}>
      {pos.replace(/_/g, ' ')}
    </span>
  );
}

const PARTICLE_POS = new Set(['particle', 'auxiliary_verb', 'conjunction']);

export function isFilterable(entry: VocabularyEntry): boolean {
  if (PARTICLE_POS.has(entry.partOfSpeech ?? '')) return true;
  if (entry.headword.length === 1 && /[\u3040-\u309F\u30A0-\u30FF]/.test(entry.headword))
    return true;
  return false;
}
