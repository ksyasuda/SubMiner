import type { ResolveContext } from '../context';
import { isObject } from '../shared';

const LEGACY_KEYS = new Set([
  'wordField',
  'audioField',
  'imageField',
  'sentenceField',
  'miscInfoField',
  'miscInfoPattern',
  'generateAudio',
  'generateImage',
  'imageType',
  'imageFormat',
  'imageQuality',
  'imageMaxWidth',
  'imageMaxHeight',
  'animatedFps',
  'animatedMaxWidth',
  'animatedMaxHeight',
  'animatedCrf',
  'syncAnimatedImageToWordAudio',
  'audioPadding',
  'fallbackDuration',
  'maxMediaDuration',
  'overwriteAudio',
  'overwriteImage',
  'mediaInsertMode',
  'highlightWord',
  'notificationType',
  'autoUpdateNewCards',
]);

export function initializeAnkiConnectResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  const {
    knownWords: _knownWordsConfigFromAnkiConnect,
    nPlusOne: _nPlusOneConfigFromAnkiConnect,
    ai: _ankiAiConfig,
    ...ankiConnectWithoutKnownWordsOrNPlusOne
  } = ankiConnect;
  const ankiConnectWithoutLegacy = Object.fromEntries(
    Object.entries(ankiConnectWithoutKnownWordsOrNPlusOne).filter(([key]) => !LEGACY_KEYS.has(key)),
  );

  context.resolved.ankiConnect = {
    ...context.resolved.ankiConnect,
    ...(isObject(ankiConnectWithoutLegacy)
      ? (ankiConnectWithoutLegacy as Partial<(typeof context.resolved)['ankiConnect']>)
      : {}),
    fields: {
      ...context.resolved.ankiConnect.fields,
    },
    media: {
      ...context.resolved.ankiConnect.media,
    },
    knownWords: {
      ...context.resolved.ankiConnect.knownWords,
    },
    behavior: {
      ...context.resolved.ankiConnect.behavior,
    },
    proxy: {
      ...context.resolved.ankiConnect.proxy,
    },
    metadata: {
      ...context.resolved.ankiConnect.metadata,
    },
    isLapis: {
      ...context.resolved.ankiConnect.isLapis,
    },
    isKiku: {
      ...context.resolved.ankiConnect.isKiku,
      ...(isObject(ankiConnect.isKiku)
        ? (ankiConnect.isKiku as (typeof context.resolved)['ankiConnect']['isKiku'])
        : {}),
    },
    lapisKiku: {
      ...context.resolved.ankiConnect.lapisKiku,
    },
  };
}
