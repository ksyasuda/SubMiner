import type { ResolveContext } from '../context';
import { asBoolean, asString, isObject } from '../shared';

export function applyLapisResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  if (isObject(ankiConnect.isLapis)) {
    const lapisEnabled = asBoolean(ankiConnect.isLapis.enabled);
    if (lapisEnabled !== undefined) {
      context.resolved.ankiConnect.isLapis.enabled = lapisEnabled;
    } else if (ankiConnect.isLapis.enabled !== undefined) {
      context.warn(
        'ankiConnect.isLapis.enabled',
        ankiConnect.isLapis.enabled,
        context.resolved.ankiConnect.isLapis.enabled,
        'Expected boolean.',
      );
    }

    const sentenceCardModel = asString(ankiConnect.isLapis.sentenceCardModel);
    if (sentenceCardModel !== undefined) {
      context.resolved.ankiConnect.isLapis.sentenceCardModel = sentenceCardModel;
    } else if (ankiConnect.isLapis.sentenceCardModel !== undefined) {
      context.warn(
        'ankiConnect.isLapis.sentenceCardModel',
        ankiConnect.isLapis.sentenceCardModel,
        context.resolved.ankiConnect.isLapis.sentenceCardModel,
        'Expected string.',
      );
    }

    if (ankiConnect.isLapis.sentenceCardSentenceField !== undefined) {
      context.warn(
        'ankiConnect.isLapis.sentenceCardSentenceField',
        ankiConnect.isLapis.sentenceCardSentenceField,
        'Sentence',
        'Deprecated key; sentence-card sentence field is fixed to Sentence.',
      );
    }

    if (ankiConnect.isLapis.sentenceCardAudioField !== undefined) {
      context.warn(
        'ankiConnect.isLapis.sentenceCardAudioField',
        ankiConnect.isLapis.sentenceCardAudioField,
        'SentenceAudio',
        'Deprecated key; sentence-card audio field is fixed to SentenceAudio.',
      );
    }
  } else if (ankiConnect.isLapis !== undefined) {
    context.warn(
      'ankiConnect.isLapis',
      ankiConnect.isLapis,
      context.resolved.ankiConnect.isLapis,
      'Expected object.',
    );
  }
}
