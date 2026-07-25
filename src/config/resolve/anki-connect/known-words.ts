import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asBoolean, asColor, asNumber, asString, isObject } from '../shared';
import { hasOwn } from './shared';

export function applyAnkiKnownWordsResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
  behavior: Record<string, unknown>,
): void {
  const knownWordsConfig = isObject(ankiConnect.knownWords) ? ankiConnect.knownWords : {};
  const nPlusOneConfig = isObject(ankiConnect.nPlusOne) ? ankiConnect.nPlusOne : {};

  const knownWordsHighlightEnabled = asBoolean(knownWordsConfig.highlightEnabled);
  if (knownWordsHighlightEnabled !== undefined) {
    context.resolved.ankiConnect.knownWords.highlightEnabled = knownWordsHighlightEnabled;
  } else if (hasOwn(knownWordsConfig, 'highlightEnabled')) {
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
    } else if (hasOwn(behavior, 'nPlusOneHighlightEnabled')) {
      context.warn(
        'ankiConnect.behavior.nPlusOneHighlightEnabled',
        behavior.nPlusOneHighlightEnabled,
        DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled,
        'Expected boolean.',
      );
      context.resolved.ankiConnect.knownWords.highlightEnabled =
        DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled;
    } else {
      context.resolved.ankiConnect.knownWords.highlightEnabled =
        DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled;
    }
  }

  const knownWordsMaturityEnabled = asBoolean(knownWordsConfig.maturityEnabled);
  if (knownWordsMaturityEnabled !== undefined) {
    context.resolved.ankiConnect.knownWords.maturityEnabled = knownWordsMaturityEnabled;
  } else if (hasOwn(knownWordsConfig, 'maturityEnabled')) {
    context.warn(
      'ankiConnect.knownWords.maturityEnabled',
      knownWordsConfig.maturityEnabled,
      context.resolved.ankiConnect.knownWords.maturityEnabled,
      'Expected boolean.',
    );
    context.resolved.ankiConnect.knownWords.maturityEnabled =
      DEFAULT_CONFIG.ankiConnect.knownWords.maturityEnabled;
  } else {
    context.resolved.ankiConnect.knownWords.maturityEnabled =
      DEFAULT_CONFIG.ankiConnect.knownWords.maturityEnabled;
  }

  const knownWordsMatureThresholdDays = asNumber(knownWordsConfig.matureThresholdDays);
  const hasValidMatureThresholdDays =
    knownWordsMatureThresholdDays !== undefined &&
    Number.isInteger(knownWordsMatureThresholdDays) &&
    knownWordsMatureThresholdDays >= 1;
  if (hasOwn(knownWordsConfig, 'matureThresholdDays')) {
    if (hasValidMatureThresholdDays) {
      context.resolved.ankiConnect.knownWords.matureThresholdDays = knownWordsMatureThresholdDays;
    } else {
      context.warn(
        'ankiConnect.knownWords.matureThresholdDays',
        knownWordsConfig.matureThresholdDays,
        DEFAULT_CONFIG.ankiConnect.knownWords.matureThresholdDays,
        'Expected an integer of at least 1.',
      );
      context.resolved.ankiConnect.knownWords.matureThresholdDays =
        DEFAULT_CONFIG.ankiConnect.knownWords.matureThresholdDays;
    }
  } else {
    context.resolved.ankiConnect.knownWords.matureThresholdDays =
      DEFAULT_CONFIG.ankiConnect.knownWords.matureThresholdDays;
  }

  const knownWordsRefreshMinutes = asNumber(knownWordsConfig.refreshMinutes);
  const hasValidKnownWordsRefreshMinutes =
    knownWordsRefreshMinutes !== undefined &&
    Number.isInteger(knownWordsRefreshMinutes) &&
    knownWordsRefreshMinutes > 0;
  if (hasOwn(knownWordsConfig, 'refreshMinutes')) {
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
  } else if (hasOwn(behavior, 'nPlusOneRefreshMinutes')) {
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
  if (hasOwn(nPlusOneConfig, 'minSentenceWords')) {
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
  } else if (hasOwn(knownWordsConfig, 'matchMode')) {
    context.warn(
      'ankiConnect.knownWords.matchMode',
      knownWordsConfig.matchMode,
      DEFAULT_CONFIG.ankiConnect.knownWords.matchMode,
      "Expected 'headword' or 'surface'.",
    );
    context.resolved.ankiConnect.knownWords.matchMode =
      DEFAULT_CONFIG.ankiConnect.knownWords.matchMode;
  } else if (hasOwn(behavior, 'nPlusOneMatchMode')) {
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

  const defaultFields = [DEFAULT_CONFIG.ankiConnect.fields.word, 'Word', 'Reading', 'Word Reading'];
  const knownWordsDecks = knownWordsConfig.decks;
  if (isObject(knownWordsDecks)) {
    const resolved: Record<string, string[]> = {};
    for (const [deck, fields] of Object.entries(knownWordsDecks)) {
      const deckName = deck.trim();
      if (!deckName) continue;
      if (Array.isArray(fields) && fields.every((field) => typeof field === 'string')) {
        resolved[deckName] = fields
          .map((field) => field.trim())
          .filter((field) => field.length > 0);
      } else {
        context.warn(
          `ankiConnect.knownWords.decks["${deckName}"]`,
          fields,
          defaultFields,
          'Expected an array of field name strings.',
        );
        resolved[deckName] = defaultFields;
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
      resolved[deck] = defaultFields;
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

  const rawSubtitleStyle = isObject(context.src.subtitleStyle) ? context.src.subtitleStyle : {};
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
}
