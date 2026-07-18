import type { ResolveContext } from '../context';
import { asBoolean, asNumber, asString } from '../shared';
import { asNotificationType, hasOwn } from './shared';

export function applyAnkiLegacyResolution(
  context: ResolveContext,
  legacy: Record<string, unknown>,
  behavior: Record<string, unknown>,
  fields: Record<string, unknown>,
  media: Record<string, unknown>,
  metadata: Record<string, unknown>,
): void {
  const asIntegerInRange = (value: unknown, min: number, max: number): number | undefined => {
    const parsed = asNumber(value);
    if (parsed === undefined || !Number.isInteger(parsed) || parsed < min || parsed > max) {
      return undefined;
    }
    return parsed;
  };
  const asPositiveInteger = (value: unknown): number | undefined => {
    const parsed = asNumber(value);
    if (parsed === undefined || !Number.isInteger(parsed) || parsed <= 0) {
      return undefined;
    }
    return parsed;
  };
  const asPositiveNumber = (value: unknown): number | undefined => {
    const parsed = asNumber(value);
    if (parsed === undefined || parsed <= 0) {
      return undefined;
    }
    return parsed;
  };
  const asNonNegativeNumber = (value: unknown): number | undefined => {
    const parsed = asNumber(value);
    if (parsed === undefined || parsed < 0) {
      return undefined;
    }
    return parsed;
  };
  const asImageType = (value: unknown): 'static' | 'avif' | undefined => {
    return value === 'static' || value === 'avif' ? value : undefined;
  };
  const asImageFormat = (value: unknown): 'jpg' | 'png' | 'webp' | undefined => {
    return value === 'jpg' || value === 'png' || value === 'webp' ? value : undefined;
  };
  const asMediaInsertMode = (value: unknown): 'append' | 'prepend' | undefined => {
    return value === 'append' || value === 'prepend' ? value : undefined;
  };
  const mapLegacy = <T>(
    key: string,
    parse: (value: unknown) => T | undefined,
    apply: (value: T) => void,
    fallback: unknown,
    message: string,
  ): void => {
    const value = legacy[key];
    if (value === undefined) return;
    const parsed = parse(value);
    if (parsed === undefined) {
      context.warn(`ankiConnect.${key}`, value, fallback, message);
      return;
    }
    apply(parsed);
  };

  if (!hasOwn(fields, 'audio')) {
    mapLegacy(
      'audioField',
      asString,
      (value) => {
        context.resolved.ankiConnect.fields.audio = value;
      },
      context.resolved.ankiConnect.fields.audio,
      'Expected string.',
    );
  }
  if (!hasOwn(fields, 'word')) {
    mapLegacy(
      'wordField',
      asString,
      (value) => {
        context.resolved.ankiConnect.fields.word = value;
      },
      context.resolved.ankiConnect.fields.word,
      'Expected string.',
    );
  }
  if (!hasOwn(fields, 'image')) {
    mapLegacy(
      'imageField',
      asString,
      (value) => {
        context.resolved.ankiConnect.fields.image = value;
      },
      context.resolved.ankiConnect.fields.image,
      'Expected string.',
    );
  }
  if (!hasOwn(fields, 'sentence')) {
    mapLegacy(
      'sentenceField',
      asString,
      (value) => {
        context.resolved.ankiConnect.fields.sentence = value;
      },
      context.resolved.ankiConnect.fields.sentence,
      'Expected string.',
    );
  }
  if (!hasOwn(fields, 'miscInfo')) {
    mapLegacy(
      'miscInfoField',
      asString,
      (value) => {
        context.resolved.ankiConnect.fields.miscInfo = value;
      },
      context.resolved.ankiConnect.fields.miscInfo,
      'Expected string.',
    );
  }
  if (!hasOwn(metadata, 'pattern')) {
    mapLegacy(
      'miscInfoPattern',
      asString,
      (value) => {
        context.resolved.ankiConnect.metadata.pattern = value;
      },
      context.resolved.ankiConnect.metadata.pattern,
      'Expected string.',
    );
  }
  if (!hasOwn(media, 'generateAudio')) {
    mapLegacy(
      'generateAudio',
      asBoolean,
      (value) => {
        context.resolved.ankiConnect.media.generateAudio = value;
      },
      context.resolved.ankiConnect.media.generateAudio,
      'Expected boolean.',
    );
  }
  if (!hasOwn(media, 'generateImage')) {
    mapLegacy(
      'generateImage',
      asBoolean,
      (value) => {
        context.resolved.ankiConnect.media.generateImage = value;
      },
      context.resolved.ankiConnect.media.generateImage,
      'Expected boolean.',
    );
  }
  if (!hasOwn(media, 'imageType')) {
    mapLegacy(
      'imageType',
      asImageType,
      (value) => {
        context.resolved.ankiConnect.media.imageType = value;
      },
      context.resolved.ankiConnect.media.imageType,
      "Expected 'static' or 'avif'.",
    );
  }
  if (!hasOwn(media, 'imageFormat')) {
    mapLegacy(
      'imageFormat',
      asImageFormat,
      (value) => {
        context.resolved.ankiConnect.media.imageFormat = value;
      },
      context.resolved.ankiConnect.media.imageFormat,
      "Expected 'jpg', 'png', or 'webp'.",
    );
  }
  if (!hasOwn(media, 'imageQuality')) {
    mapLegacy(
      'imageQuality',
      (value) => asIntegerInRange(value, 1, 100),
      (value) => {
        context.resolved.ankiConnect.media.imageQuality = value;
      },
      context.resolved.ankiConnect.media.imageQuality,
      'Expected integer between 1 and 100.',
    );
  }
  if (!hasOwn(media, 'imageMaxWidth')) {
    mapLegacy(
      'imageMaxWidth',
      asPositiveInteger,
      (value) => {
        context.resolved.ankiConnect.media.imageMaxWidth = value;
      },
      context.resolved.ankiConnect.media.imageMaxWidth,
      'Expected positive integer.',
    );
  }
  if (!hasOwn(media, 'imageMaxHeight')) {
    mapLegacy(
      'imageMaxHeight',
      asPositiveInteger,
      (value) => {
        context.resolved.ankiConnect.media.imageMaxHeight = value;
      },
      context.resolved.ankiConnect.media.imageMaxHeight,
      'Expected positive integer.',
    );
  }
  if (!hasOwn(media, 'animatedFps')) {
    mapLegacy(
      'animatedFps',
      (value) => asIntegerInRange(value, 1, 60),
      (value) => {
        context.resolved.ankiConnect.media.animatedFps = value;
      },
      context.resolved.ankiConnect.media.animatedFps,
      'Expected integer between 1 and 60.',
    );
  }
  if (!hasOwn(media, 'animatedMaxWidth')) {
    mapLegacy(
      'animatedMaxWidth',
      asPositiveInteger,
      (value) => {
        context.resolved.ankiConnect.media.animatedMaxWidth = value;
      },
      context.resolved.ankiConnect.media.animatedMaxWidth,
      'Expected positive integer.',
    );
  }
  if (!hasOwn(media, 'animatedMaxHeight')) {
    mapLegacy(
      'animatedMaxHeight',
      asPositiveInteger,
      (value) => {
        context.resolved.ankiConnect.media.animatedMaxHeight = value;
      },
      context.resolved.ankiConnect.media.animatedMaxHeight,
      'Expected positive integer.',
    );
  }
  if (!hasOwn(media, 'animatedCrf')) {
    mapLegacy(
      'animatedCrf',
      (value) => asIntegerInRange(value, 0, 63),
      (value) => {
        context.resolved.ankiConnect.media.animatedCrf = value;
      },
      context.resolved.ankiConnect.media.animatedCrf,
      'Expected integer between 0 and 63.',
    );
  }
  if (!hasOwn(media, 'syncAnimatedImageToWordAudio')) {
    mapLegacy(
      'syncAnimatedImageToWordAudio',
      asBoolean,
      (value) => {
        context.resolved.ankiConnect.media.syncAnimatedImageToWordAudio = value;
      },
      context.resolved.ankiConnect.media.syncAnimatedImageToWordAudio,
      'Expected boolean.',
    );
  }
  if (!hasOwn(media, 'audioPadding')) {
    mapLegacy(
      'audioPadding',
      asNonNegativeNumber,
      (value) => {
        context.resolved.ankiConnect.media.audioPadding = value;
      },
      context.resolved.ankiConnect.media.audioPadding,
      'Expected non-negative number.',
    );
  }
  if (!hasOwn(media, 'fallbackDuration')) {
    mapLegacy(
      'fallbackDuration',
      asPositiveNumber,
      (value) => {
        context.resolved.ankiConnect.media.fallbackDuration = value;
      },
      context.resolved.ankiConnect.media.fallbackDuration,
      'Expected positive number.',
    );
  }
  if (!hasOwn(media, 'maxMediaDuration')) {
    mapLegacy(
      'maxMediaDuration',
      asNonNegativeNumber,
      (value) => {
        context.resolved.ankiConnect.media.maxMediaDuration = value;
      },
      context.resolved.ankiConnect.media.maxMediaDuration,
      'Expected non-negative number.',
    );
  }
  if (!hasOwn(behavior, 'overwriteAudio')) {
    mapLegacy(
      'overwriteAudio',
      asBoolean,
      (value) => {
        context.resolved.ankiConnect.behavior.overwriteAudio = value;
      },
      context.resolved.ankiConnect.behavior.overwriteAudio,
      'Expected boolean.',
    );
  }
  if (!hasOwn(behavior, 'overwriteImage')) {
    mapLegacy(
      'overwriteImage',
      asBoolean,
      (value) => {
        context.resolved.ankiConnect.behavior.overwriteImage = value;
      },
      context.resolved.ankiConnect.behavior.overwriteImage,
      'Expected boolean.',
    );
  }
  if (!hasOwn(behavior, 'mediaInsertMode')) {
    mapLegacy(
      'mediaInsertMode',
      asMediaInsertMode,
      (value) => {
        context.resolved.ankiConnect.behavior.mediaInsertMode = value;
      },
      context.resolved.ankiConnect.behavior.mediaInsertMode,
      "Expected 'append' or 'prepend'.",
    );
  }
  if (!hasOwn(behavior, 'highlightWord')) {
    mapLegacy(
      'highlightWord',
      asBoolean,
      (value) => {
        context.resolved.ankiConnect.behavior.highlightWord = value;
      },
      context.resolved.ankiConnect.behavior.highlightWord,
      'Expected boolean.',
    );
  }
  if (!hasOwn(behavior, 'notificationType')) {
    mapLegacy(
      'notificationType',
      asNotificationType,
      (value) => {
        context.resolved.ankiConnect.behavior.notificationType = value;
      },
      context.resolved.ankiConnect.behavior.notificationType,
      "Expected 'overlay', 'system', 'both', 'none', 'osd', or 'osd-system'.",
    );
  }
  if (!hasOwn(behavior, 'autoUpdateNewCards')) {
    mapLegacy(
      'autoUpdateNewCards',
      asBoolean,
      (value) => {
        context.resolved.ankiConnect.behavior.autoUpdateNewCards = value;
      },
      context.resolved.ankiConnect.behavior.autoUpdateNewCards,
      'Expected boolean.',
    );
  }
}
