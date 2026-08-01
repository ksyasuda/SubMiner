import type { CardKind, WordCardKind } from '../types/anki';

/**
 * Kiku/Lapis note types decide which card a note generates from mutually exclusive
 * `Is...Card` flag fields. Setting one always means clearing the others.
 */
export const CARD_KIND_FLAG_FIELDS: Record<CardKind, string> = {
  'word-and-sentence': 'IsWordAndSentenceCard',
  click: 'IsClickCard',
  sentence: 'IsSentenceCard',
  audio: 'IsAudioCard',
};

export const WORD_CARD_KINDS: readonly WordCardKind[] = [
  'word-and-sentence',
  'click',
  'sentence',
  'audio',
  'none',
];

export const DEFAULT_WORD_CARD_KIND: WordCardKind = 'word-and-sentence';

/**
 * Card kinds SubMiner marks on its own initiative (word cards). They are only applied
 * when the note type actually carries the matching flag field, so plain note types keep
 * their fields untouched.
 */
const IMPLICIT_CARD_KINDS = new Set<CardKind>(['word-and-sentence', 'click']);

export function isWordCardKind(value: unknown): value is WordCardKind {
  return typeof value === 'string' && WORD_CARD_KINDS.includes(value as WordCardKind);
}

export function resolveWordCardKindSetting(value: unknown): WordCardKind {
  return isWordCardKind(value) ? value : DEFAULT_WORD_CARD_KIND;
}

/**
 * Flags `cardKind` on the note and clears every other card-kind flag it has, so the note
 * never ends up claiming to be two kinds of card at once.
 */
export function applyCardKindFlagFields(
  updatedFields: Record<string, string>,
  cardKind: CardKind,
  resolveFieldName: (preferredName: string) => string | null,
): void {
  const targetFlag = resolveFieldName(CARD_KIND_FLAG_FIELDS[cardKind]);
  if (!targetFlag && IMPLICIT_CARD_KINDS.has(cardKind)) {
    return;
  }
  if (targetFlag) {
    updatedFields[targetFlag] = 'x';
  }

  for (const [kind, flagName] of Object.entries(CARD_KIND_FLAG_FIELDS)) {
    if (kind === cardKind) continue;
    const resolved = resolveFieldName(flagName);
    if (resolved && resolved !== targetFlag) {
      updatedFields[resolved] = '';
    }
  }
}
