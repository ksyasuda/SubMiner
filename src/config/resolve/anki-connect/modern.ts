import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asBoolean, asNumber, asString, isObject } from '../shared';
import { asNotificationType, hasOwn } from './shared';

function asIntegerInRange(value: unknown, min: number, max: number): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && Number.isInteger(parsed) && parsed >= min && parsed <= max
    ? parsed
    : undefined;
}

function asNonNegativeInteger(value: unknown): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

function asPositiveNumber(value: unknown): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && parsed > 0 ? parsed : undefined;
}

function asNonNegativeNumber(value: unknown): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && parsed >= 0 ? parsed : undefined;
}

function applyModernValue<T>(
  context: ResolveContext,
  source: Record<string, unknown>,
  key: string,
  path: string,
  parse: (value: unknown) => T | undefined,
  fallback: T,
  apply: (value: T) => void,
  message: string,
): void {
  if (!hasOwn(source, key)) return;
  const raw = source[key];
  const parsed = parse(raw);
  if (parsed === undefined) {
    apply(fallback);
    context.warn(path, raw, fallback, message);
    return;
  }
  apply(parsed);
}

function applyModernFieldsResolution(
  context: ResolveContext,
  fields: Record<string, unknown>,
): void {
  for (const key of ['word', 'audio', 'image', 'sentence', 'miscInfo', 'translation'] as const) {
    applyModernValue(
      context,
      fields,
      key,
      `ankiConnect.fields.${key}`,
      asString,
      DEFAULT_CONFIG.ankiConnect.fields[key],
      (value) => {
        context.resolved.ankiConnect.fields[key] = value;
      },
      'Expected string.',
    );
  }
}

function applyModernMediaResolution(context: ResolveContext, media: Record<string, unknown>): void {
  for (const key of [
    'generateAudio',
    'generateImage',
    'syncAnimatedImageToWordAudio',
    'normalizeAudio',
    'mirrorMpvVolume',
  ] as const) {
    applyModernValue(
      context,
      media,
      key,
      `ankiConnect.media.${key}`,
      asBoolean,
      DEFAULT_CONFIG.ankiConnect.media[key],
      (value) => {
        context.resolved.ankiConnect.media[key] = value;
      },
      'Expected boolean.',
    );
  }

  applyModernValue(
    context,
    media,
    'imageType',
    'ankiConnect.media.imageType',
    (value) => (value === 'static' || value === 'avif' ? value : undefined),
    DEFAULT_CONFIG.ankiConnect.media.imageType,
    (value) => {
      context.resolved.ankiConnect.media.imageType = value;
    },
    "Expected 'static' or 'avif'.",
  );
  applyModernValue(
    context,
    media,
    'imageFormat',
    'ankiConnect.media.imageFormat',
    (value) => (value === 'jpg' || value === 'png' || value === 'webp' ? value : undefined),
    DEFAULT_CONFIG.ankiConnect.media.imageFormat,
    (value) => {
      context.resolved.ankiConnect.media.imageFormat = value;
    },
    "Expected 'jpg', 'png', or 'webp'.",
  );
  applyModernValue(
    context,
    media,
    'imageQuality',
    'ankiConnect.media.imageQuality',
    (value) => asIntegerInRange(value, 1, 100),
    DEFAULT_CONFIG.ankiConnect.media.imageQuality,
    (value) => {
      context.resolved.ankiConnect.media.imageQuality = value;
    },
    'Expected integer between 1 and 100.',
  );

  for (const key of [
    'imageMaxWidth',
    'imageMaxHeight',
    'animatedMaxWidth',
    'animatedMaxHeight',
  ] as const) {
    applyModernValue(
      context,
      media,
      key,
      `ankiConnect.media.${key}`,
      asNonNegativeInteger,
      DEFAULT_CONFIG.ankiConnect.media[key] ?? 0,
      (value) => {
        context.resolved.ankiConnect.media[key] = value;
      },
      'Expected non-negative integer.',
    );
  }

  applyModernValue(
    context,
    media,
    'animatedFps',
    'ankiConnect.media.animatedFps',
    (value) => asIntegerInRange(value, 1, 60),
    DEFAULT_CONFIG.ankiConnect.media.animatedFps,
    (value) => {
      context.resolved.ankiConnect.media.animatedFps = value;
    },
    'Expected integer between 1 and 60.',
  );
  applyModernValue(
    context,
    media,
    'animatedCrf',
    'ankiConnect.media.animatedCrf',
    (value) => asIntegerInRange(value, 0, 63),
    DEFAULT_CONFIG.ankiConnect.media.animatedCrf,
    (value) => {
      context.resolved.ankiConnect.media.animatedCrf = value;
    },
    'Expected integer between 0 and 63.',
  );
  applyModernValue(
    context,
    media,
    'audioPadding',
    'ankiConnect.media.audioPadding',
    asNonNegativeNumber,
    DEFAULT_CONFIG.ankiConnect.media.audioPadding,
    (value) => {
      context.resolved.ankiConnect.media.audioPadding = value;
    },
    'Expected non-negative number.',
  );

  for (const key of ['fallbackDuration', 'maxMediaDuration'] as const) {
    applyModernValue(
      context,
      media,
      key,
      `ankiConnect.media.${key}`,
      asPositiveNumber,
      DEFAULT_CONFIG.ankiConnect.media[key],
      (value) => {
        context.resolved.ankiConnect.media[key] = value;
      },
      'Expected positive number.',
    );
  }
}

function applyModernBehaviorResolution(
  context: ResolveContext,
  behavior: Record<string, unknown>,
): void {
  for (const key of [
    'overwriteAudio',
    'overwriteImage',
    'highlightWord',
    'autoUpdateNewCards',
  ] as const) {
    applyModernValue(
      context,
      behavior,
      key,
      `ankiConnect.behavior.${key}`,
      asBoolean,
      DEFAULT_CONFIG.ankiConnect.behavior[key],
      (value) => {
        context.resolved.ankiConnect.behavior[key] = value;
      },
      'Expected boolean.',
    );
  }

  applyModernValue(
    context,
    behavior,
    'mediaInsertMode',
    'ankiConnect.behavior.mediaInsertMode',
    (value) => (value === 'append' || value === 'prepend' ? value : undefined),
    DEFAULT_CONFIG.ankiConnect.behavior.mediaInsertMode,
    (value) => {
      context.resolved.ankiConnect.behavior.mediaInsertMode = value;
    },
    "Expected 'append' or 'prepend'.",
  );
  applyModernValue(
    context,
    behavior,
    'notificationType',
    'ankiConnect.behavior.notificationType',
    asNotificationType,
    DEFAULT_CONFIG.ankiConnect.behavior.notificationType,
    (value) => {
      context.resolved.ankiConnect.behavior.notificationType = value;
    },
    "Expected 'overlay', 'system', 'both', 'none', 'osd', or 'osd-system'.",
  );
}

function applyModernMetadataResolution(
  context: ResolveContext,
  metadata: Record<string, unknown>,
): void {
  applyModernValue(
    context,
    metadata,
    'pattern',
    'ankiConnect.metadata.pattern',
    asString,
    DEFAULT_CONFIG.ankiConnect.metadata.pattern,
    (value) => {
      context.resolved.ankiConnect.metadata.pattern = value;
    },
    'Expected string.',
  );
}

export function applyAnkiModernResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
  behavior: Record<string, unknown>,
  media: Record<string, unknown>,
): void {
  const fields = isObject(ankiConnect.fields) ? ankiConnect.fields : {};
  const metadata = isObject(ankiConnect.metadata) ? ankiConnect.metadata : {};
  applyModernFieldsResolution(context, fields);
  applyModernMediaResolution(context, media);
  applyModernBehaviorResolution(context, behavior);
  applyModernMetadataResolution(context, metadata);

  applyLapisResolution(context, ankiConnect);
  applyProxyResolution(context, ankiConnect);
  applyAiResolution(context, ankiConnect);
  applyTagsResolution(context, ankiConnect);
}

function applyLapisResolution(context: ResolveContext, ankiConnect: Record<string, unknown>): void {
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

function applyProxyResolution(context: ResolveContext, ankiConnect: Record<string, unknown>): void {
  if (isObject(ankiConnect.proxy)) {
    const proxy = ankiConnect.proxy;
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
  } else if (ankiConnect.proxy !== undefined) {
    context.warn(
      'ankiConnect.proxy',
      ankiConnect.proxy,
      context.resolved.ankiConnect.proxy,
      'Expected object.',
    );
  }
}

function applyAiResolution(context: ResolveContext, ankiConnect: Record<string, unknown>): void {
  if (isObject(ankiConnect.ai)) {
    const aiEnabled = asBoolean(ankiConnect.ai.enabled);
    if (aiEnabled !== undefined) {
      context.resolved.ankiConnect.ai.enabled = aiEnabled;
    } else if (ankiConnect.ai.enabled !== undefined) {
      context.warn(
        'ankiConnect.ai.enabled',
        ankiConnect.ai.enabled,
        context.resolved.ankiConnect.ai.enabled,
        'Expected boolean.',
      );
    }

    const aiModel = asString(ankiConnect.ai.model);
    if (aiModel !== undefined) {
      context.resolved.ankiConnect.ai.model = aiModel;
    } else if (ankiConnect.ai.model !== undefined) {
      context.warn(
        'ankiConnect.ai.model',
        ankiConnect.ai.model,
        context.resolved.ankiConnect.ai.model,
        'Expected string.',
      );
    }

    const aiSystemPrompt = asString(ankiConnect.ai.systemPrompt);
    if (aiSystemPrompt !== undefined) {
      context.resolved.ankiConnect.ai.systemPrompt = aiSystemPrompt;
    } else if (ankiConnect.ai.systemPrompt !== undefined) {
      context.warn(
        'ankiConnect.ai.systemPrompt',
        ankiConnect.ai.systemPrompt,
        context.resolved.ankiConnect.ai.systemPrompt,
        'Expected string.',
      );
    }
  } else {
    const aiEnabled = asBoolean(ankiConnect.ai);
    if (aiEnabled !== undefined) {
      context.resolved.ankiConnect.ai.enabled = aiEnabled;
    } else if (ankiConnect.ai !== undefined) {
      context.warn(
        'ankiConnect.ai',
        ankiConnect.ai,
        context.resolved.ankiConnect.ai.enabled,
        'Expected boolean or object.',
      );
    }
  }
}

function applyTagsResolution(context: ResolveContext, ankiConnect: Record<string, unknown>): void {
  if (Array.isArray(ankiConnect.tags)) {
    const normalizedTags = ankiConnect.tags
      .filter((entry): entry is string => typeof entry === 'string')
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);
    if (normalizedTags.length === ankiConnect.tags.length) {
      context.resolved.ankiConnect.tags = [...new Set(normalizedTags)];
    } else {
      context.resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
      context.warn(
        'ankiConnect.tags',
        ankiConnect.tags,
        context.resolved.ankiConnect.tags,
        'Expected an array of non-empty strings.',
      );
    }
  } else if (ankiConnect.tags !== undefined) {
    context.resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
    context.warn(
      'ankiConnect.tags',
      ankiConnect.tags,
      context.resolved.ankiConnect.tags,
      'Expected an array of strings.',
    );
  }
}

export function applyAnkiKikuResolution(context: ResolveContext): void {
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
