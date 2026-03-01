import { markNPlusOneTargets } from '../../../token-merger';
import {
  DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG,
  resolveAnnotationPos1ExclusionSet,
} from '../../../token-pos1-exclusions';
import {
  DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG,
  resolveAnnotationPos2ExclusionSet,
} from '../../../token-pos2-exclusions';
import { JlptLevel, MergedToken, NPlusOneMatchMode, PartOfSpeech } from '../../../types';
import { shouldIgnoreJlptByTerm, shouldIgnoreJlptForMecabPos1 } from '../jlpt-token-filter';

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;
const JLPT_LEVEL_LOOKUP_CACHE_LIMIT = 2048;

const jlptLevelLookupCaches = new WeakMap<
  (text: string) => JlptLevel | null,
  Map<string, JlptLevel | null>
>();

export interface AnnotationStageDeps {
  isKnownWord: (text: string) => boolean;
  knownWordMatchMode: NPlusOneMatchMode;
  getJlptLevel: (text: string) => JlptLevel | null;
}

export interface AnnotationStageOptions {
  nPlusOneEnabled?: boolean;
  jlptEnabled?: boolean;
  frequencyEnabled?: boolean;
  minSentenceWordsForNPlusOne?: number;
  pos1Exclusions?: ReadonlySet<string>;
  pos2Exclusions?: ReadonlySet<string>;
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

function normalizePos1Tag(pos1: string | undefined): string {
  return typeof pos1 === 'string' ? pos1.trim() : '';
}

function isExcludedByTagSet(normalizedTag: string, exclusions: ReadonlySet<string>): boolean {
  if (!normalizedTag) {
    return false;
  }
  const parts = normalizedTag
    .split('|')
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
  if (parts.length === 0) {
    return false;
  }
  // Composite tags like "助詞|名詞" stay eligible unless every component is excluded.
  return parts.every((part) => exclusions.has(part));
}

function resolvePos1Exclusions(options: AnnotationStageOptions): ReadonlySet<string> {
  if (options.pos1Exclusions) {
    return options.pos1Exclusions;
  }

  return resolveAnnotationPos1ExclusionSet(DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG);
}

function resolvePos2Exclusions(options: AnnotationStageOptions): ReadonlySet<string> {
  if (options.pos2Exclusions) {
    return options.pos2Exclusions;
  }

  return resolveAnnotationPos2ExclusionSet(DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG);
}

function normalizePos2Tag(pos2: string | undefined): string {
  return typeof pos2 === 'string' ? pos2.trim() : '';
}

function isFrequencyExcludedByPos(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): boolean {
  const normalizedPos1 = normalizePos1Tag(token.pos1);
  const hasPos1 = normalizedPos1.length > 0;
  if (isExcludedByTagSet(normalizedPos1, pos1Exclusions)) {
    return true;
  }

  const normalizedPos2 = normalizePos2Tag(token.pos2);
  const hasPos2 = normalizedPos2.length > 0;
  if (isExcludedByTagSet(normalizedPos2, pos2Exclusions)) {
    return true;
  }

  if (hasPos1 || hasPos2) {
    return false;
  }

  if (isLikelyFrequencyNoiseToken(token)) {
    return true;
  }

  return (
    token.partOfSpeech === PartOfSpeech.particle ||
    token.partOfSpeech === PartOfSpeech.bound_auxiliary
  );
}

function applyFrequencyMarking(
  tokens: MergedToken[],
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): MergedToken[] {
  return tokens.map((token) => {
    if (isFrequencyExcludedByPos(token, pos1Exclusions, pos2Exclusions)) {
      return { ...token, frequencyRank: undefined };
    }

    if (typeof token.frequencyRank === 'number' && Number.isFinite(token.frequencyRank)) {
      const rank = Math.max(1, Math.floor(token.frequencyRank));
      return { ...token, frequencyRank: rank };
    }

    return {
      ...token,
      frequencyRank: undefined,
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

function isTrailingSmallTsuKanaSfx(text: string): boolean {
  const normalized = normalizeJlptTextForExclusion(text);
  if (!normalized) {
    return false;
  }

  const chars = [...normalized];
  if (chars.length < 2 || chars.length > 4) {
    return false;
  }

  if (!chars.every(isKanaChar)) {
    return false;
  }

  return chars[chars.length - 1] === 'っ';
}

function isReduplicatedKanaSfx(text: string): boolean {
  const normalized = normalizeJlptTextForExclusion(text);
  if (!normalized) {
    return false;
  }

  const chars = [...normalized];
  if (chars.length < 4 || chars.length % 2 !== 0) {
    return false;
  }

  if (!chars.every(isKanaChar)) {
    return false;
  }

  const half = chars.length / 2;
  return chars.slice(0, half).join('') === chars.slice(half).join('');
}

function hasAdjacentKanaRepeat(text: string): boolean {
  const normalized = normalizeJlptTextForExclusion(text);
  if (!normalized) {
    return false;
  }

  const chars = [...normalized];
  if (!chars.every(isKanaChar)) {
    return false;
  }

  for (let i = 1; i < chars.length; i += 1) {
    if (chars[i] === chars[i - 1]) {
      return true;
    }
  }

  return false;
}

function isLikelyFrequencyNoiseToken(token: MergedToken): boolean {
  const candidates = [token.headword, token.surface].filter(
    (candidate): candidate is string => typeof candidate === 'string' && candidate.length > 0,
  );

  for (const candidate of candidates) {
    const trimmedCandidate = candidate.trim();
    if (!trimmedCandidate) {
      continue;
    }

    const normalizedCandidate = normalizeJlptTextForExclusion(trimmedCandidate);
    if (!normalizedCandidate) {
      continue;
    }

    if (shouldIgnoreJlptByTerm(trimmedCandidate) || shouldIgnoreJlptByTerm(normalizedCandidate)) {
      return true;
    }

    if (
      hasAdjacentKanaRepeat(trimmedCandidate) ||
      hasAdjacentKanaRepeat(normalizedCandidate) ||
      isReduplicatedKanaSfx(trimmedCandidate) ||
      isReduplicatedKanaSfx(normalizedCandidate) ||
      isTrailingSmallTsuKanaSfx(trimmedCandidate) ||
      isTrailingSmallTsuKanaSfx(normalizedCandidate)
    ) {
      return true;
    }
  }

  return false;
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
  const pos1Exclusions = resolvePos1Exclusions(options);
  const pos2Exclusions = resolvePos2Exclusions(options);
  const nPlusOneEnabled = options.nPlusOneEnabled !== false;
  const knownMarkedTokens = nPlusOneEnabled
    ? applyKnownWordMarking(tokens, deps.isKnownWord, deps.knownWordMatchMode)
    : tokens.map((token) => ({
        ...token,
        isKnown: false,
        isNPlusOneTarget: false,
      }));

  const frequencyEnabled = options.frequencyEnabled !== false;
  const frequencyMarkedTokens = frequencyEnabled
    ? applyFrequencyMarking(knownMarkedTokens, pos1Exclusions, pos2Exclusions)
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

  return markNPlusOneTargets(
    jlptMarkedTokens,
    sanitizedMinSentenceWordsForNPlusOne,
    pos1Exclusions,
    pos2Exclusions,
  );
}
