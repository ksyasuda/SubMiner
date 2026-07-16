import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asBoolean, asNumber, asString, isObject } from '../shared';
import { asNotificationType, hasOwn } from './shared';

export function applyAnkiModernResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
  behavior: Record<string, unknown>,
  media: Record<string, unknown>,
): void {
  if (hasOwn(media, 'mirrorMpvVolume')) {
    const parsed = asBoolean(media.mirrorMpvVolume);
    if (parsed === undefined) {
      context.resolved.ankiConnect.media.mirrorMpvVolume =
        DEFAULT_CONFIG.ankiConnect.media.mirrorMpvVolume;
      context.warn(
        'ankiConnect.media.mirrorMpvVolume',
        media.mirrorMpvVolume,
        context.resolved.ankiConnect.media.mirrorMpvVolume,
        'Expected boolean.',
      );
    } else {
      context.resolved.ankiConnect.media.mirrorMpvVolume = parsed;
    }
  }

  if (hasOwn(behavior, 'notificationType')) {
    const parsed = asNotificationType(behavior.notificationType);
    if (parsed === undefined) {
      context.resolved.ankiConnect.behavior.notificationType =
        DEFAULT_CONFIG.ankiConnect.behavior.notificationType;
      context.warn(
        'ankiConnect.behavior.notificationType',
        behavior.notificationType,
        context.resolved.ankiConnect.behavior.notificationType,
        "Expected 'overlay', 'system', 'both', 'none', 'osd', or 'osd-system'.",
      );
    } else {
      context.resolved.ankiConnect.behavior.notificationType = parsed;
    }
  }

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
