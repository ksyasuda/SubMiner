import { markNPlusOneTargets } from '../../../token-merger';
import {
  FrequencyDictionaryLookup,
  JlptLevel,
  MergedToken,
  NPlusOneMatchMode,
  PartOfSpeech,
} from '../../../types';
import { shouldIgnoreJlptByTerm, shouldIgnoreJlptForMecabPos1 } from '../jlpt-token-filter';

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;
const JLPT_LEVEL_LOOKUP_CACHE_LIMIT = 2048;
const FREQUENCY_RANK_LOOKUP_CACHE_LIMIT = 2048;

const jlptLevelLookupCaches = new WeakMap<
  (text: string) => JlptLevel | null,
  Map<string, JlptLevel | null>
>();
const frequencyRankLookupCaches = new WeakMap<
  FrequencyDictionaryLookup,
  Map<string, number | null>
>();

export interface AnnotationStageDeps {
  isKnownWord: (text: string) => boolean;
  knownWordMatchMode: NPlusOneMatchMode;
  getJlptLevel: (text: string) => JlptLevel | null;
  getFrequencyRank?: FrequencyDictionaryLookup;
}

export interface AnnotationStageOptions {
  nPlusOneEnabled?: boolean;
  jlptEnabled?: boolean;
  frequencyEnabled?: boolean;
  minSentenceWordsForNPlusOne?: number;
}

function resolveKnownWordText(
  surface: string,
  headword: string,
  matchMode: NPlusOneMatchMode,
): string {
  return matchMode === 'surface' ? surface : headword;
}

function applyKnownWordMarking(
  tokens: MergedToken[],
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): MergedToken[] {
  return tokens.map((token) => {
    const matchText = resolveKnownWordText(token.surface, token.headword, knownWordMatchMode);

    return {
      ...token,
      isKnown: token.isKnown || (matchText ? isKnownWord(matchText) : false),
    };
  });
}

function normalizeFrequencyLookupText(rawText: string): string {
  return rawText.trim().toLowerCase();
}

function getCachedFrequencyRank(
  lookupText: string,
  getFrequencyRank: FrequencyDictionaryLookup,
): number | null {
  const normalizedText = normalizeFrequencyLookupText(lookupText);
  if (!normalizedText) {
    return null;
  }

  let cache = frequencyRankLookupCaches.get(getFrequencyRank);
  if (!cache) {
    cache = new Map<string, number | null>();
    frequencyRankLookupCaches.set(getFrequencyRank, cache);
  }

  if (cache.has(normalizedText)) {
    return cache.get(normalizedText) ?? null;
  }

  let rank: number | null;
  try {
    rank = getFrequencyRank(normalizedText);
  } catch {
    rank = null;
  }
  if (rank !== null) {
    if (!Number.isFinite(rank) || rank <= 0) {
      rank = null;
    }
  }

  cache.set(normalizedText, rank);
  while (cache.size > FREQUENCY_RANK_LOOKUP_CACHE_LIMIT) {
    const firstKey = cache.keys().next().value;
    if (firstKey !== undefined) {
      cache.delete(firstKey);
    }
  }

  return rank;
}

function resolveFrequencyLookupText(token: MergedToken): string {
  if (token.headword && token.headword.length > 0) {
    return token.headword;
  }
  if (token.reading && token.reading.length > 0) {
    return token.reading;
  }
  return token.surface;
}

function getFrequencyLookupTextCandidates(token: MergedToken): string[] {
  const lookupText = resolveFrequencyLookupText(token).trim();
  return lookupText ? [lookupText] : [];
}

function isFrequencyExcludedByPos(token: MergedToken): boolean {
  if (
    token.partOfSpeech === PartOfSpeech.particle ||
    token.partOfSpeech === PartOfSpeech.bound_auxiliary
  ) {
    return true;
  }

  return token.pos1 === '助詞' || token.pos1 === '助動詞';
}

function applyFrequencyMarking(
  tokens: MergedToken[],
  getFrequencyRank: FrequencyDictionaryLookup,
): MergedToken[] {
  return tokens.map((token) => {
    if (isFrequencyExcludedByPos(token)) {
      return { ...token, frequencyRank: undefined };
    }

    if (typeof token.frequencyRank === 'number' && Number.isFinite(token.frequencyRank)) {
      const rank = Math.max(1, Math.floor(token.frequencyRank));
      return { ...token, frequencyRank: rank };
    }

    const lookupTexts = getFrequencyLookupTextCandidates(token);
    if (lookupTexts.length === 0) {
      return { ...token, frequencyRank: undefined };
    }

    let bestRank: number | null = null;
    for (const lookupText of lookupTexts) {
      const rank = getCachedFrequencyRank(lookupText, getFrequencyRank);
      if (rank === null) {
        continue;
      }
      if (bestRank === null || rank < bestRank) {
        bestRank = rank;
      }
    }

    return {
      ...token,
      frequencyRank: bestRank ?? undefined,
    };
  });
}

function getCachedJlptLevel(
  lookupText: string,
  getJlptLevel: (text: string) => JlptLevel | null,
): JlptLevel | null {
  const normalizedText = lookupText.trim();
  if (!normalizedText) {
    return null;
  }

  let cache = jlptLevelLookupCaches.get(getJlptLevel);
  if (!cache) {
    cache = new Map<string, JlptLevel | null>();
    jlptLevelLookupCaches.set(getJlptLevel, cache);
  }

  if (cache.has(normalizedText)) {
    return cache.get(normalizedText) ?? null;
  }

  let level: JlptLevel | null;
  try {
    level = getJlptLevel(normalizedText);
  } catch {
    level = null;
  }

  cache.set(normalizedText, level);
  while (cache.size > JLPT_LEVEL_LOOKUP_CACHE_LIMIT) {
    const firstKey = cache.keys().next().value;
    if (firstKey !== undefined) {
      cache.delete(firstKey);
    }
  }

  return level;
}

function resolveJlptLookupText(token: MergedToken): string {
  if (token.headword && token.headword.length > 0) {
    return token.headword;
  }
  if (token.reading && token.reading.length > 0) {
    return token.reading;
  }
  return token.surface;
}

function normalizeJlptTextForExclusion(text: string): string {
  const raw = text.trim();
  if (!raw) {
    return '';
  }

  let normalized = '';
  for (const char of raw) {
    const code = char.codePointAt(0);
    if (code === undefined) {
      continue;
    }

    if (code >= KATAKANA_CODEPOINT_START && code <= KATAKANA_CODEPOINT_END) {
      normalized += String.fromCodePoint(code - KATAKANA_TO_HIRAGANA_OFFSET);
      continue;
    }

    normalized += char;
  }

  return normalized;
}

function isKanaChar(char: string): boolean {
  const code = char.codePointAt(0);
  if (code === undefined) {
    return false;
  }

  return (
    (code >= 0x3041 && code <= 0x3096) ||
    (code >= 0x309b && code <= 0x309f) ||
    (code >= 0x30a0 && code <= 0x30fa) ||
    (code >= 0x30fd && code <= 0x30ff)
  );
}

function isRepeatedKanaSfx(text: string): boolean {
  const normalized = text.trim();
  if (!normalized) {
    return false;
  }

  const chars = [...normalized];
  if (!chars.every(isKanaChar)) {
    return false;
  }

  const counts = new Map<string, number>();
  let hasAdjacentRepeat = false;

  for (let i = 0; i < chars.length; i += 1) {
    const char = chars[i]!;
    counts.set(char, (counts.get(char) ?? 0) + 1);
    if (i > 0 && chars[i] === chars[i - 1]) {
      hasAdjacentRepeat = true;
    }
  }

  const topCount = Math.max(...counts.values());
  if (chars.length <= 2) {
    return hasAdjacentRepeat || topCount >= 2;
  }

  if (hasAdjacentRepeat) {
    return true;
  }

  return topCount >= Math.ceil(chars.length / 2);
}

function isJlptEligibleToken(token: MergedToken): boolean {
  if (token.pos1 && shouldIgnoreJlptForMecabPos1(token.pos1)) {
    return false;
  }

  const candidates = [
    resolveJlptLookupText(token),
    token.surface,
    token.reading,
    token.headword,
  ].filter(
    (candidate): candidate is string => typeof candidate === 'string' && candidate.length > 0,
  );

  for (const candidate of candidates) {
    const normalizedCandidate = normalizeJlptTextForExclusion(candidate);
    if (!normalizedCandidate) {
      continue;
    }

    const trimmedCandidate = candidate.trim();
    if (shouldIgnoreJlptByTerm(trimmedCandidate) || shouldIgnoreJlptByTerm(normalizedCandidate)) {
      return false;
    }

    if (isRepeatedKanaSfx(candidate) || isRepeatedKanaSfx(normalizedCandidate)) {
      return false;
    }
  }

  return true;
}

function applyJlptMarking(
  tokens: MergedToken[],
  getJlptLevel: (text: string) => JlptLevel | null,
): MergedToken[] {
  return tokens.map((token) => {
    if (!isJlptEligibleToken(token)) {
      return { ...token, jlptLevel: undefined };
    }

    const primaryLevel = getCachedJlptLevel(resolveJlptLookupText(token), getJlptLevel);
    const fallbackLevel =
      primaryLevel === null ? getCachedJlptLevel(token.surface, getJlptLevel) : null;

    return {
      ...token,
      jlptLevel: primaryLevel ?? fallbackLevel ?? token.jlptLevel,
    };
  });
}

export function annotateTokens(
  tokens: MergedToken[],
  deps: AnnotationStageDeps,
  options: AnnotationStageOptions = {},
): MergedToken[] {
  const nPlusOneEnabled = options.nPlusOneEnabled !== false;
  const knownMarkedTokens = nPlusOneEnabled
    ? applyKnownWordMarking(tokens, deps.isKnownWord, deps.knownWordMatchMode)
    : tokens.map((token) => ({
        ...token,
        isKnown: false,
        isNPlusOneTarget: false,
      }));

  const frequencyEnabled = options.frequencyEnabled !== false;
  const frequencyMarkedTokens =
    frequencyEnabled && deps.getFrequencyRank
      ? applyFrequencyMarking(knownMarkedTokens, deps.getFrequencyRank)
      : frequencyEnabled
        ? knownMarkedTokens.map((token) => ({
            ...token,
            frequencyRank:
              typeof token.frequencyRank === 'number' && Number.isFinite(token.frequencyRank)
                ? Math.max(1, Math.floor(token.frequencyRank))
                : undefined,
          }))
      : knownMarkedTokens.map((token) => ({
          ...token,
          frequencyRank: undefined,
        }));

  const jlptEnabled = options.jlptEnabled !== false;
  const jlptMarkedTokens = jlptEnabled
    ? applyJlptMarking(frequencyMarkedTokens, deps.getJlptLevel)
    : frequencyMarkedTokens.map((token) => ({
        ...token,
        jlptLevel: undefined,
      }));

  if (!nPlusOneEnabled) {
    return jlptMarkedTokens.map((token) => ({
      ...token,
      isKnown: false,
      isNPlusOneTarget: false,
    }));
  }

  const minSentenceWordsForNPlusOne = options.minSentenceWordsForNPlusOne;
  const sanitizedMinSentenceWordsForNPlusOne =
    minSentenceWordsForNPlusOne !== undefined &&
    Number.isInteger(minSentenceWordsForNPlusOne) &&
    minSentenceWordsForNPlusOne > 0
      ? minSentenceWordsForNPlusOne
      : 3;

  return markNPlusOneTargets(jlptMarkedTokens, sanitizedMinSentenceWordsForNPlusOne);
}
