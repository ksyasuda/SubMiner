import { DEFAULT_CONFIG } from '../definitions';
import type { ResolveContext } from './context';
import { asBoolean, asColor, asNumber, asString, isObject } from './shared';

export function applyAnkiConnectResolution(context: ResolveContext): void {
  if (!isObject(context.src.ankiConnect)) {
    return;
  }

  const ac = context.src.ankiConnect;
  const behavior = isObject(ac.behavior) ? (ac.behavior as Record<string, unknown>) : {};
  const fields = isObject(ac.fields) ? (ac.fields as Record<string, unknown>) : {};
  const media = isObject(ac.media) ? (ac.media as Record<string, unknown>) : {};
  const metadata = isObject(ac.metadata) ? (ac.metadata as Record<string, unknown>) : {};
  const proxy = isObject(ac.proxy) ? (ac.proxy as Record<string, unknown>) : {};
  const legacyKeys = new Set([
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

  const {
    knownWords: _knownWordsConfigFromAnkiConnect,
    nPlusOne: _nPlusOneConfigFromAnkiConnect,
    ai: _ankiAiConfig,
    ...ankiConnectWithoutKnownWordsOrNPlusOne
  } = ac as Record<string, unknown>;
  const ankiConnectWithoutLegacy = Object.fromEntries(
    Object.entries(ankiConnectWithoutKnownWordsOrNPlusOne).filter(([key]) => !legacyKeys.has(key)),
  );

  context.resolved.ankiConnect = {
    ...context.resolved.ankiConnect,
    ...(isObject(ankiConnectWithoutLegacy)
      ? (ankiConnectWithoutLegacy as Partial<(typeof context.resolved)['ankiConnect']>)
      : {}),
    fields: {
      ...context.resolved.ankiConnect.fields,
      ...(isObject(ac.fields)
        ? (ac.fields as (typeof context.resolved)['ankiConnect']['fields'])
        : {}),
    },
    media: {
      ...context.resolved.ankiConnect.media,
      ...(isObject(ac.media)
        ? (ac.media as (typeof context.resolved)['ankiConnect']['media'])
        : {}),
    },
    knownWords: {
      ...context.resolved.ankiConnect.knownWords,
    },
    behavior: {
      ...context.resolved.ankiConnect.behavior,
      ...(isObject(ac.behavior)
        ? (ac.behavior as (typeof context.resolved)['ankiConnect']['behavior'])
        : {}),
    },
    proxy: {
      ...context.resolved.ankiConnect.proxy,
    },
    metadata: {
      ...context.resolved.ankiConnect.metadata,
      ...(isObject(ac.metadata)
        ? (ac.metadata as (typeof context.resolved)['ankiConnect']['metadata'])
        : {}),
    },
    isLapis: {
      ...context.resolved.ankiConnect.isLapis,
    },
    isKiku: {
      ...context.resolved.ankiConnect.isKiku,
      ...(isObject(ac.isKiku)
        ? (ac.isKiku as (typeof context.resolved)['ankiConnect']['isKiku'])
        : {}),
    },
  };

  if (isObject(ac.isLapis)) {
    const lapisEnabled = asBoolean(ac.isLapis.enabled);
    if (lapisEnabled !== undefined) {
      context.resolved.ankiConnect.isLapis.enabled = lapisEnabled;
    } else if (ac.isLapis.enabled !== undefined) {
      context.warn(
        'ankiConnect.isLapis.enabled',
        ac.isLapis.enabled,
        context.resolved.ankiConnect.isLapis.enabled,
        'Expected boolean.',
      );
    }

    const sentenceCardModel = asString(ac.isLapis.sentenceCardModel);
    if (sentenceCardModel !== undefined) {
      context.resolved.ankiConnect.isLapis.sentenceCardModel = sentenceCardModel;
    } else if (ac.isLapis.sentenceCardModel !== undefined) {
      context.warn(
        'ankiConnect.isLapis.sentenceCardModel',
        ac.isLapis.sentenceCardModel,
        context.resolved.ankiConnect.isLapis.sentenceCardModel,
        'Expected string.',
      );
    }

    if (ac.isLapis.sentenceCardSentenceField !== undefined) {
      context.warn(
        'ankiConnect.isLapis.sentenceCardSentenceField',
        ac.isLapis.sentenceCardSentenceField,
        'Sentence',
        'Deprecated key; sentence-card sentence field is fixed to Sentence.',
      );
    }

    if (ac.isLapis.sentenceCardAudioField !== undefined) {
      context.warn(
        'ankiConnect.isLapis.sentenceCardAudioField',
        ac.isLapis.sentenceCardAudioField,
        'SentenceAudio',
        'Deprecated key; sentence-card audio field is fixed to SentenceAudio.',
      );
    }
  } else if (ac.isLapis !== undefined) {
    context.warn(
      'ankiConnect.isLapis',
      ac.isLapis,
      context.resolved.ankiConnect.isLapis,
      'Expected object.',
    );
  }

  if (isObject(ac.proxy)) {
    const proxyEnabled = asBoolean(proxy.enabled);
    if (proxyEnabled !== undefined) {
      context.resolved.ankiConnect.proxy.enabled = proxyEnabled;
    } else if (proxy.enabled !== undefined) {
      context.warn(
        'ankiConnect.proxy.enabled',
        proxy.enabled,
        context.resolved.ankiConnect.proxy.enabled,
        'Expected boolean.',
      );
    }

    const proxyHost = asString(proxy.host);
    if (proxyHost !== undefined && proxyHost.trim().length > 0) {
      context.resolved.ankiConnect.proxy.host = proxyHost.trim();
    } else if (proxy.host !== undefined) {
      context.warn(
        'ankiConnect.proxy.host',
        proxy.host,
        context.resolved.ankiConnect.proxy.host,
        'Expected non-empty string.',
      );
    }

    const proxyUpstreamUrl = asString(proxy.upstreamUrl);
    if (proxyUpstreamUrl !== undefined && proxyUpstreamUrl.trim().length > 0) {
      context.resolved.ankiConnect.proxy.upstreamUrl = proxyUpstreamUrl.trim();
    } else if (proxy.upstreamUrl !== undefined) {
      context.warn(
        'ankiConnect.proxy.upstreamUrl',
        proxy.upstreamUrl,
        context.resolved.ankiConnect.proxy.upstreamUrl,
        'Expected non-empty string.',
      );
    }

    const proxyPort = asNumber(proxy.port);
    if (
      proxyPort !== undefined &&
      Number.isInteger(proxyPort) &&
      proxyPort >= 1 &&
      proxyPort <= 65535
    ) {
      context.resolved.ankiConnect.proxy.port = proxyPort;
    } else if (proxy.port !== undefined) {
      context.warn(
        'ankiConnect.proxy.port',
        proxy.port,
        context.resolved.ankiConnect.proxy.port,
        'Expected integer between 1 and 65535.',
      );
    }
  } else if (ac.proxy !== undefined) {
    context.warn(
      'ankiConnect.proxy',
      ac.proxy,
      context.resolved.ankiConnect.proxy,
      'Expected object.',
    );
  }

  if (isObject(ac.ai)) {
    const aiEnabled = asBoolean(ac.ai.enabled);
    if (aiEnabled !== undefined) {
      context.resolved.ankiConnect.ai.enabled = aiEnabled;
    } else if (ac.ai.enabled !== undefined) {
      context.warn(
        'ankiConnect.ai.enabled',
        ac.ai.enabled,
        context.resolved.ankiConnect.ai.enabled,
        'Expected boolean.',
      );
    }

    const aiModel = asString(ac.ai.model);
    if (aiModel !== undefined) {
      context.resolved.ankiConnect.ai.model = aiModel;
    } else if (ac.ai.model !== undefined) {
      context.warn(
        'ankiConnect.ai.model',
        ac.ai.model,
        context.resolved.ankiConnect.ai.model,
        'Expected string.',
      );
    }

    const aiSystemPrompt = asString(ac.ai.systemPrompt);
    if (aiSystemPrompt !== undefined) {
      context.resolved.ankiConnect.ai.systemPrompt = aiSystemPrompt;
    } else if (ac.ai.systemPrompt !== undefined) {
      context.warn(
        'ankiConnect.ai.systemPrompt',
        ac.ai.systemPrompt,
        context.resolved.ankiConnect.ai.systemPrompt,
        'Expected string.',
      );
    }
  } else {
    const aiEnabled = asBoolean(ac.ai);
    if (aiEnabled !== undefined) {
      context.resolved.ankiConnect.ai.enabled = aiEnabled;
    } else if (ac.ai !== undefined) {
      context.warn(
        'ankiConnect.ai',
        ac.ai,
        context.resolved.ankiConnect.ai.enabled,
        'Expected boolean or object.',
      );
    }
  }

  if (Array.isArray(ac.tags)) {
    const normalizedTags = ac.tags
      .filter((entry): entry is string => typeof entry === 'string')
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);
    if (normalizedTags.length === ac.tags.length) {
      context.resolved.ankiConnect.tags = [...new Set(normalizedTags)];
    } else {
      context.resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
      context.warn(
        'ankiConnect.tags',
        ac.tags,
        context.resolved.ankiConnect.tags,
        'Expected an array of non-empty strings.',
      );
    }
  } else if (ac.tags !== undefined) {
    context.resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
    context.warn(
      'ankiConnect.tags',
      ac.tags,
      context.resolved.ankiConnect.tags,
      'Expected an array of strings.',
    );
  }

  const legacy = ac as Record<string, unknown>;
  const hasOwn = (obj: Record<string, unknown>, key: string): boolean =>
    Object.prototype.hasOwnProperty.call(obj, key);
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
  const asNotificationType = (value: unknown): 'osd' | 'system' | 'both' | 'none' | undefined => {
    return value === 'osd' || value === 'system' || value === 'both' || value === 'none'
      ? value
      : undefined;
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
      "Expected 'osd', 'system', 'both', or 'none'.",
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

  const knownWordsConfig = isObject(ac.knownWords)
    ? (ac.knownWords as Record<string, unknown>)
    : {};
  const nPlusOneConfig = isObject(ac.nPlusOne) ? (ac.nPlusOne as Record<string, unknown>) : {};

  const knownWordsHighlightEnabled = asBoolean(knownWordsConfig.highlightEnabled);
  if (knownWordsHighlightEnabled !== undefined) {
    context.resolved.ankiConnect.knownWords.highlightEnabled = knownWordsHighlightEnabled;
  } else if (knownWordsConfig.highlightEnabled !== undefined) {
    context.warn(
      'ankiConnect.knownWords.highlightEnabled',
      knownWordsConfig.highlightEnabled,
      context.resolved.ankiConnect.knownWords.highlightEnabled,
      'Expected boolean.',
    );
    context.resolved.ankiConnect.knownWords.highlightEnabled =
      DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled;
  } else {
    const legacyBehaviorNPlusOneHighlightEnabled = asBoolean(behavior.nPlusOneHighlightEnabled);
    if (legacyBehaviorNPlusOneHighlightEnabled !== undefined) {
      context.resolved.ankiConnect.knownWords.highlightEnabled =
        legacyBehaviorNPlusOneHighlightEnabled;
      context.warn(
        'ankiConnect.behavior.nPlusOneHighlightEnabled',
        behavior.nPlusOneHighlightEnabled,
        DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled,
        'Legacy key is deprecated; use ankiConnect.knownWords.highlightEnabled',
      );
    } else {
      context.resolved.ankiConnect.knownWords.highlightEnabled =
        DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled;
    }
  }

  const knownWordsRefreshMinutes = asNumber(knownWordsConfig.refreshMinutes);
  const hasValidKnownWordsRefreshMinutes =
    knownWordsRefreshMinutes !== undefined &&
    Number.isInteger(knownWordsRefreshMinutes) &&
    knownWordsRefreshMinutes > 0;
  if (knownWordsRefreshMinutes !== undefined) {
    if (hasValidKnownWordsRefreshMinutes) {
      context.resolved.ankiConnect.knownWords.refreshMinutes = knownWordsRefreshMinutes;
    } else {
      context.warn(
        'ankiConnect.knownWords.refreshMinutes',
        knownWordsConfig.refreshMinutes,
        context.resolved.ankiConnect.knownWords.refreshMinutes,
        'Expected a positive integer.',
      );
      context.resolved.ankiConnect.knownWords.refreshMinutes =
        DEFAULT_CONFIG.ankiConnect.knownWords.refreshMinutes;
    }
  } else if (asNumber(behavior.nPlusOneRefreshMinutes) !== undefined) {
    const legacyBehaviorNPlusOneRefreshMinutes = asNumber(behavior.nPlusOneRefreshMinutes);
    const hasValidLegacyRefreshMinutes =
      legacyBehaviorNPlusOneRefreshMinutes !== undefined &&
      Number.isInteger(legacyBehaviorNPlusOneRefreshMinutes) &&
      legacyBehaviorNPlusOneRefreshMinutes > 0;
    if (hasValidLegacyRefreshMinutes) {
      context.resolved.ankiConnect.knownWords.refreshMinutes = legacyBehaviorNPlusOneRefreshMinutes;
      context.warn(
        'ankiConnect.behavior.nPlusOneRefreshMinutes',
        behavior.nPlusOneRefreshMinutes,
        DEFAULT_CONFIG.ankiConnect.knownWords.refreshMinutes,
        'Legacy key is deprecated; use ankiConnect.knownWords.refreshMinutes',
      );
    } else {
      context.warn(
        'ankiConnect.behavior.nPlusOneRefreshMinutes',
        behavior.nPlusOneRefreshMinutes,
        context.resolved.ankiConnect.knownWords.refreshMinutes,
        'Expected a positive integer.',
      );
      context.resolved.ankiConnect.knownWords.refreshMinutes =
        DEFAULT_CONFIG.ankiConnect.knownWords.refreshMinutes;
    }
  } else {
    context.resolved.ankiConnect.knownWords.refreshMinutes =
      DEFAULT_CONFIG.ankiConnect.knownWords.refreshMinutes;
  }

  const knownWordsAddMinedWordsImmediately = asBoolean(knownWordsConfig.addMinedWordsImmediately);
  if (knownWordsAddMinedWordsImmediately !== undefined) {
    context.resolved.ankiConnect.knownWords.addMinedWordsImmediately =
      knownWordsAddMinedWordsImmediately;
  } else if (knownWordsConfig.addMinedWordsImmediately !== undefined) {
    context.warn(
      'ankiConnect.knownWords.addMinedWordsImmediately',
      knownWordsConfig.addMinedWordsImmediately,
      context.resolved.ankiConnect.knownWords.addMinedWordsImmediately,
      'Expected boolean.',
    );
    context.resolved.ankiConnect.knownWords.addMinedWordsImmediately =
      DEFAULT_CONFIG.ankiConnect.knownWords.addMinedWordsImmediately;
  } else {
    context.resolved.ankiConnect.knownWords.addMinedWordsImmediately =
      DEFAULT_CONFIG.ankiConnect.knownWords.addMinedWordsImmediately;
  }

  const nPlusOneEnabled = asBoolean(nPlusOneConfig.enabled);
  if (nPlusOneEnabled !== undefined) {
    context.resolved.ankiConnect.nPlusOne.enabled = nPlusOneEnabled;
  } else if (nPlusOneConfig.enabled !== undefined) {
    context.warn(
      'ankiConnect.nPlusOne.enabled',
      nPlusOneConfig.enabled,
      context.resolved.ankiConnect.nPlusOne.enabled,
      'Expected boolean.',
    );
    context.resolved.ankiConnect.nPlusOne.enabled = DEFAULT_CONFIG.ankiConnect.nPlusOne.enabled;
  } else {
    context.resolved.ankiConnect.nPlusOne.enabled = DEFAULT_CONFIG.ankiConnect.nPlusOne.enabled;
  }

  const nPlusOneMinSentenceWords = asNumber(nPlusOneConfig.minSentenceWords);
  const hasValidNPlusOneMinSentenceWords =
    nPlusOneMinSentenceWords !== undefined &&
    Number.isInteger(nPlusOneMinSentenceWords) &&
    nPlusOneMinSentenceWords > 0;
  if (nPlusOneMinSentenceWords !== undefined) {
    if (hasValidNPlusOneMinSentenceWords) {
      context.resolved.ankiConnect.nPlusOne.minSentenceWords = nPlusOneMinSentenceWords;
    } else {
      context.warn(
        'ankiConnect.nPlusOne.minSentenceWords',
        nPlusOneConfig.minSentenceWords,
        context.resolved.ankiConnect.nPlusOne.minSentenceWords,
        'Expected a positive integer.',
      );
      context.resolved.ankiConnect.nPlusOne.minSentenceWords =
        DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords;
    }
  } else {
    context.resolved.ankiConnect.nPlusOne.minSentenceWords =
      DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords;
  }

  const knownWordsMatchMode = asString(knownWordsConfig.matchMode);
  const legacyBehaviorNPlusOneMatchMode = asString(behavior.nPlusOneMatchMode);
  const hasValidKnownWordsMatchMode =
    knownWordsMatchMode === 'headword' || knownWordsMatchMode === 'surface';
  const hasValidLegacyMatchMode =
    legacyBehaviorNPlusOneMatchMode === 'headword' || legacyBehaviorNPlusOneMatchMode === 'surface';
  if (hasValidKnownWordsMatchMode) {
    context.resolved.ankiConnect.knownWords.matchMode = knownWordsMatchMode;
  } else if (knownWordsMatchMode !== undefined) {
    context.warn(
      'ankiConnect.knownWords.matchMode',
      knownWordsConfig.matchMode,
      DEFAULT_CONFIG.ankiConnect.knownWords.matchMode,
      "Expected 'headword' or 'surface'.",
    );
    context.resolved.ankiConnect.knownWords.matchMode =
      DEFAULT_CONFIG.ankiConnect.knownWords.matchMode;
  } else if (legacyBehaviorNPlusOneMatchMode !== undefined) {
    if (hasValidLegacyMatchMode) {
      context.resolved.ankiConnect.knownWords.matchMode = legacyBehaviorNPlusOneMatchMode;
      context.warn(
        'ankiConnect.behavior.nPlusOneMatchMode',
        behavior.nPlusOneMatchMode,
        DEFAULT_CONFIG.ankiConnect.knownWords.matchMode,
        'Legacy key is deprecated; use ankiConnect.knownWords.matchMode',
      );
    } else {
      context.warn(
        'ankiConnect.behavior.nPlusOneMatchMode',
        behavior.nPlusOneMatchMode,
        context.resolved.ankiConnect.knownWords.matchMode,
        "Expected 'headword' or 'surface'.",
      );
      context.resolved.ankiConnect.knownWords.matchMode =
        DEFAULT_CONFIG.ankiConnect.knownWords.matchMode;
    }
  } else {
    context.resolved.ankiConnect.knownWords.matchMode =
      DEFAULT_CONFIG.ankiConnect.knownWords.matchMode;
  }

  const DEFAULT_FIELDS = [
    DEFAULT_CONFIG.ankiConnect.fields.word,
    'Word',
    'Reading',
    'Word Reading',
  ];
  const knownWordsDecks = knownWordsConfig.decks;
  if (isObject(knownWordsDecks)) {
    const resolved: Record<string, string[]> = {};
    for (const [deck, fields] of Object.entries(knownWordsDecks as Record<string, unknown>)) {
      const deckName = deck.trim();
      if (!deckName) continue;
      if (Array.isArray(fields) && fields.every((f) => typeof f === 'string')) {
        resolved[deckName] = (fields as string[]).map((f) => f.trim()).filter((f) => f.length > 0);
      } else {
        context.warn(
          `ankiConnect.knownWords.decks["${deckName}"]`,
          fields,
          DEFAULT_FIELDS,
          'Expected an array of field name strings.',
        );
        resolved[deckName] = DEFAULT_FIELDS;
      }
    }
    context.resolved.ankiConnect.knownWords.decks = resolved;
  } else if (Array.isArray(knownWordsDecks)) {
    const normalized = knownWordsDecks
      .filter((entry): entry is string => typeof entry === 'string')
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);
    const resolved: Record<string, string[]> = {};
    for (const deck of new Set(normalized)) {
      resolved[deck] = DEFAULT_FIELDS;
    }
    context.resolved.ankiConnect.knownWords.decks = resolved;
    if (normalized.length > 0) {
      context.warn(
        'ankiConnect.knownWords.decks',
        knownWordsDecks,
        resolved,
        'Legacy array format is deprecated; use object format: { "Deck Name": ["Field1", "Field2"] }',
      );
    }
  } else if (knownWordsDecks !== undefined) {
    context.warn(
      'ankiConnect.knownWords.decks',
      knownWordsDecks,
      context.resolved.ankiConnect.knownWords.decks,
      'Expected an object mapping deck names to field arrays.',
    );
  }

  const rawSubtitleStyle = isObject(context.src.subtitleStyle)
    ? (context.src.subtitleStyle as Record<string, unknown>)
    : {};
  const hasCanonicalKnownWordColor = rawSubtitleStyle.knownWordColor !== undefined;

  const knownWordsColor = asColor(knownWordsConfig.color);
  if (knownWordsColor !== undefined) {
    if (!hasCanonicalKnownWordColor) {
      context.resolved.subtitleStyle.knownWordColor = knownWordsColor;
    }
    context.warn(
      'ankiConnect.knownWords.color',
      knownWordsConfig.color,
      context.resolved.subtitleStyle.knownWordColor,
      'Legacy key is deprecated; use subtitleStyle.knownWordColor',
    );
  } else if (knownWordsConfig.color !== undefined) {
    context.warn(
      'ankiConnect.knownWords.color',
      knownWordsConfig.color,
      context.resolved.subtitleStyle.knownWordColor,
      'Expected a hex color value.',
    );
  }

  if (
    context.resolved.ankiConnect.isKiku.fieldGrouping !== 'auto' &&
    context.resolved.ankiConnect.isKiku.fieldGrouping !== 'manual' &&
    context.resolved.ankiConnect.isKiku.fieldGrouping !== 'disabled'
  ) {
    context.warn(
      'ankiConnect.isKiku.fieldGrouping',
      context.resolved.ankiConnect.isKiku.fieldGrouping,
      DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping,
      'Expected auto, manual, or disabled.',
    );
    context.resolved.ankiConnect.isKiku.fieldGrouping =
      DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping;
  }
}
