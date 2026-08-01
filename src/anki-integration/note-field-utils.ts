import type { CardKind, WordCardKind } from '../types/anki';
import { createLogger } from '../logger';
import {
  CARD_KIND_FLAG_FIELDS,
  DEFAULT_WORD_CARD_KIND,
  resolveWordCardKindSetting,
} from './card-kinds';

const log = createLogger('anki').child('integration.note-fields');

export interface NoteFieldValueInfo {
  fields: Record<string, { value: string }>;
}

export function getNoteFieldValue(
  noteInfo: NoteFieldValueInfo,
  preferredName: string,
): string | null {
  const resolvedFieldName = Object.keys(noteInfo.fields).find(
    (fieldName) => fieldName.toLowerCase() === preferredName.toLowerCase(),
  );
  return resolvedFieldName ? (noteInfo.fields[resolvedFieldName]?.value ?? '') : null;
}

export function hasNoteFieldValue(noteInfo: NoteFieldValueInfo, preferredName: string): boolean {
  return (getNoteFieldValue(noteInfo, preferredName) ?? '').trim().length > 0;
}

/** Flags set only by an explicit mine action; a note carrying one is not a word card. */
const EXPLICIT_CARD_FLAG_FIELDS = [CARD_KIND_FLAG_FIELDS.sentence, CARD_KIND_FLAG_FIELDS.audio];

const warnedMissingFlagFields = new Set<CardKind>();

function warnMissingFlagFieldOnce(wordCardKind: CardKind, flagField: string): void {
  if (wordCardKind === DEFAULT_WORD_CARD_KIND || warnedMissingFlagFields.has(wordCardKind)) {
    // The default kind is also the fallback for plain note types, so its absence is expected.
    return;
  }
  warnedMissingFlagFields.add(wordCardKind);
  log.warn(
    `Word card type "${wordCardKind}" is configured but the note has no ${flagField} field; leaving card type flags unchanged.`,
  );
}

/**
 * Card kind to flag when SubMiner fills a word card's sentence, or null to leave the
 * card-kind flags alone. Kiku/Lapis only: other note types have no such fields.
 */
export function resolveWordCardKind(
  noteInfo: NoteFieldValueInfo,
  sentenceCardConfig: {
    lapisEnabled: boolean;
    kikuEnabled: boolean;
    wordCardKind?: WordCardKind;
  },
): CardKind | null {
  if (!sentenceCardConfig.lapisEnabled && !sentenceCardConfig.kikuEnabled) {
    return null;
  }

  const wordCardKind = resolveWordCardKindSetting(sentenceCardConfig.wordCardKind);
  if (wordCardKind === 'none') {
    return null;
  }

  const flagField = CARD_KIND_FLAG_FIELDS[wordCardKind];
  const flagValue = getNoteFieldValue(noteInfo, flagField);
  if (flagValue === null) {
    // Note type has no flag field for the configured kind.
    warnMissingFlagFieldOnce(wordCardKind, flagField);
    return null;
  }
  if (flagValue.trim().length > 0) {
    return wordCardKind;
  }

  const alreadyExplicitCard = EXPLICIT_CARD_FLAG_FIELDS.some(
    (fieldName) =>
      fieldName.toLowerCase() !== flagField.toLowerCase() && hasNoteFieldValue(noteInfo, fieldName),
  );
  return alreadyExplicitCard ? null : wordCardKind;
}
