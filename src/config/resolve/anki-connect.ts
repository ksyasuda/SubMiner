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
    nPlusOne: _nPlusOneConfigFromAnkiConnect,
    ai: _ankiAiConfig,
    ...ankiConnectWithoutNPlusOne
  } = ac as Record<string, unknown>;
  const ankiConnectWithoutLegacy = Object.fromEntries(
    Object.entries(ankiConnectWithoutNPlusOne).filter(([key]) => !legacyKeys.has(key)),
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

  const nPlusOneConfig = isObject(ac.nPlusOne) ? (ac.nPlusOne as Record<string, unknown>) : {};

  const nPlusOneHighlightEnabled = asBoolean(nPlusOneConfig.highlightEnabled);
  if (nPlusOneHighlightEnabled !== undefined) {
    context.resolved.ankiConnect.nPlusOne.highlightEnabled = nPlusOneHighlightEnabled;
  } else if (nPlusOneConfig.highlightEnabled !== undefined) {
    context.warn(
      'ankiConnect.nPlusOne.highlightEnabled',
      nPlusOneConfig.highlightEnabled,
      context.resolved.ankiConnect.nPlusOne.highlightEnabled,
      'Expected boolean.',
    );
    context.resolved.ankiConnect.nPlusOne.highlightEnabled =
      DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled;
  } else {
    const legacyNPlusOneHighlightEnabled = asBoolean(behavior.nPlusOneHighlightEnabled);
    if (legacyNPlusOneHighlightEnabled !== undefined) {
      context.resolved.ankiConnect.nPlusOne.highlightEnabled = legacyNPlusOneHighlightEnabled;
      context.warn(
        'ankiConnect.behavior.nPlusOneHighlightEnabled',
        behavior.nPlusOneHighlightEnabled,
        DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled,
        'Legacy key is deprecated; use ankiConnect.nPlusOne.highlightEnabled',
      );
    } else {
      context.resolved.ankiConnect.nPlusOne.highlightEnabled =
        DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled;
    }
  }

  const nPlusOneRefreshMinutes = asNumber(nPlusOneConfig.refreshMinutes);
  const hasValidNPlusOneRefreshMinutes =
    nPlusOneRefreshMinutes !== undefined &&
    Number.isInteger(nPlusOneRefreshMinutes) &&
    nPlusOneRefreshMinutes > 0;
  if (nPlusOneRefreshMinutes !== undefined) {
    if (hasValidNPlusOneRefreshMinutes) {
      context.resolved.ankiConnect.nPlusOne.refreshMinutes = nPlusOneRefreshMinutes;
    } else {
      context.warn(
        'ankiConnect.nPlusOne.refreshMinutes',
        nPlusOneConfig.refreshMinutes,
        context.resolved.ankiConnect.nPlusOne.refreshMinutes,
        'Expected a positive integer.',
      );
      context.resolved.ankiConnect.nPlusOne.refreshMinutes =
        DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes;
    }
  } else if (asNumber(behavior.nPlusOneRefreshMinutes) !== undefined) {
    const legacyNPlusOneRefreshMinutes = asNumber(behavior.nPlusOneRefreshMinutes);
    const hasValidLegacyRefreshMinutes =
      legacyNPlusOneRefreshMinutes !== undefined &&
      Number.isInteger(legacyNPlusOneRefreshMinutes) &&
      legacyNPlusOneRefreshMinutes > 0;
    if (hasValidLegacyRefreshMinutes) {
      context.resolved.ankiConnect.nPlusOne.refreshMinutes = legacyNPlusOneRefreshMinutes;
      context.warn(
        'ankiConnect.behavior.nPlusOneRefreshMinutes',
        behavior.nPlusOneRefreshMinutes,
        DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes,
        'Legacy key is deprecated; use ankiConnect.nPlusOne.refreshMinutes',
      );
    } else {
      context.warn(
        'ankiConnect.behavior.nPlusOneRefreshMinutes',
        behavior.nPlusOneRefreshMinutes,
        context.resolved.ankiConnect.nPlusOne.refreshMinutes,
        'Expected a positive integer.',
      );
      context.resolved.ankiConnect.nPlusOne.refreshMinutes =
        DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes;
    }
  } else {
    context.resolved.ankiConnect.nPlusOne.refreshMinutes =
      DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes;
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

  const nPlusOneMatchMode = asString(nPlusOneConfig.matchMode);
  const legacyNPlusOneMatchMode = asString(behavior.nPlusOneMatchMode);
  const hasValidNPlusOneMatchMode =
    nPlusOneMatchMode === 'headword' || nPlusOneMatchMode === 'surface';
  const hasValidLegacyMatchMode =
    legacyNPlusOneMatchMode === 'headword' || legacyNPlusOneMatchMode === 'surface';
  if (hasValidNPlusOneMatchMode) {
    context.resolved.ankiConnect.nPlusOne.matchMode = nPlusOneMatchMode;
  } else if (nPlusOneMatchMode !== undefined) {
    context.warn(
      'ankiConnect.nPlusOne.matchMode',
      nPlusOneConfig.matchMode,
      DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
      "Expected 'headword' or 'surface'.",
    );
    context.resolved.ankiConnect.nPlusOne.matchMode = DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode;
  } else if (legacyNPlusOneMatchMode !== undefined) {
    if (hasValidLegacyMatchMode) {
      context.resolved.ankiConnect.nPlusOne.matchMode = legacyNPlusOneMatchMode;
      context.warn(
        'ankiConnect.behavior.nPlusOneMatchMode',
        behavior.nPlusOneMatchMode,
        DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
        'Legacy key is deprecated; use ankiConnect.nPlusOne.matchMode',
      );
    } else {
      context.warn(
        'ankiConnect.behavior.nPlusOneMatchMode',
        behavior.nPlusOneMatchMode,
        context.resolved.ankiConnect.nPlusOne.matchMode,
        "Expected 'headword' or 'surface'.",
      );
      context.resolved.ankiConnect.nPlusOne.matchMode =
        DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode;
    }
  } else {
    context.resolved.ankiConnect.nPlusOne.matchMode = DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode;
  }

  const nPlusOneDecks = nPlusOneConfig.decks;
  if (Array.isArray(nPlusOneDecks)) {
    const normalizedDecks = nPlusOneDecks
      .filter((entry): entry is string => typeof entry === 'string')
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);

    if (normalizedDecks.length === nPlusOneDecks.length) {
      context.resolved.ankiConnect.nPlusOne.decks = [...new Set(normalizedDecks)];
    } else if (nPlusOneDecks.length > 0) {
      context.warn(
        'ankiConnect.nPlusOne.decks',
        nPlusOneDecks,
        context.resolved.ankiConnect.nPlusOne.decks,
        'Expected an array of strings.',
      );
    } else {
      context.resolved.ankiConnect.nPlusOne.decks = [];
    }
  } else if (nPlusOneDecks !== undefined) {
    context.warn(
      'ankiConnect.nPlusOne.decks',
      nPlusOneDecks,
      context.resolved.ankiConnect.nPlusOne.decks,
      'Expected an array of strings.',
    );
    context.resolved.ankiConnect.nPlusOne.decks = [];
  }

  const nPlusOneHighlightColor = asColor(nPlusOneConfig.nPlusOne);
  if (nPlusOneHighlightColor !== undefined) {
    context.resolved.ankiConnect.nPlusOne.nPlusOne = nPlusOneHighlightColor;
  } else if (nPlusOneConfig.nPlusOne !== undefined) {
    context.warn(
      'ankiConnect.nPlusOne.nPlusOne',
      nPlusOneConfig.nPlusOne,
      context.resolved.ankiConnect.nPlusOne.nPlusOne,
      'Expected a hex color value.',
    );
    context.resolved.ankiConnect.nPlusOne.nPlusOne = DEFAULT_CONFIG.ankiConnect.nPlusOne.nPlusOne;
  }

  const nPlusOneKnownWordColor = asColor(nPlusOneConfig.knownWord);
  if (nPlusOneKnownWordColor !== undefined) {
    context.resolved.ankiConnect.nPlusOne.knownWord = nPlusOneKnownWordColor;
  } else if (nPlusOneConfig.knownWord !== undefined) {
    context.warn(
      'ankiConnect.nPlusOne.knownWord',
      nPlusOneConfig.knownWord,
      context.resolved.ankiConnect.nPlusOne.knownWord,
      'Expected a hex color value.',
    );
    context.resolved.ankiConnect.nPlusOne.knownWord = DEFAULT_CONFIG.ankiConnect.nPlusOne.knownWord;
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
