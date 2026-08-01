import { isWordCardKind, WORD_CARD_KINDS } from '../../../anki-integration/card-kinds';
import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { isObject } from '../shared';

export function applyAnkiLapisKikuResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  if (!isObject(ankiConnect.lapisKiku)) {
    if (ankiConnect.lapisKiku !== undefined) {
      context.warn(
        'ankiConnect.lapisKiku',
        ankiConnect.lapisKiku,
        context.resolved.ankiConnect.lapisKiku,
        'Expected object.',
      );
    }
    return;
  }

  const wordCardKind = ankiConnect.lapisKiku.wordCardKind;
  if (wordCardKind === undefined) {
    return;
  }
  if (isWordCardKind(wordCardKind)) {
    context.resolved.ankiConnect.lapisKiku.wordCardKind = wordCardKind;
    return;
  }

  context.warn(
    'ankiConnect.lapisKiku.wordCardKind',
    wordCardKind,
    DEFAULT_CONFIG.ankiConnect.lapisKiku.wordCardKind,
    `Expected one of ${WORD_CARD_KINDS.join(', ')}.`,
  );
  context.resolved.ankiConnect.lapisKiku.wordCardKind =
    DEFAULT_CONFIG.ankiConnect.lapisKiku.wordCardKind;
}
