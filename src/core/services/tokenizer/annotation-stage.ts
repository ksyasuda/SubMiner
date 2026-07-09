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
import {
  isKanjiNonIndependentNounToken,
  shouldExcludeTokenFromSubtitleAnnotations as sharedShouldExcludeTokenFromSubtitleAnnotations,
  stripSubtitleAnnotationMetadata as sharedStripSubtitleAnnotationMetadata,
} from './subtitle-annotation-filter';

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;
const JLPT_LEVEL_LOOKUP_CACHE_LIMIT = 2048;

const jlptLevelLookupCaches = new WeakMap<
  (text: string) => JlptLevel | null,
  Map<string, JlptLevel | null>
>();

export interface AnnotationStageDeps {
  isKnownWord: (text: string, reading?: string) => boolean;
  knownWordMatchMode: NPlusOneMatchMode;
  getJlptLevel: (text: string) => JlptLevel | null;
}

export interface AnnotationStageOptions {
  knownWordsEnabled?: boolean;
  nPlusOneEnabled?: boolean;
  nameMatchEnabled?: boolean;
  jlptEnabled?: boolean;
  frequencyEnabled?: boolean;
  minSentenceWordsForNPlusOne?: number;
  pos1Exclusions?: ReadonlySet<string>;
  pos2Exclusions?: ReadonlySet<string>;
  sourceText?: string;
}

function resolveKnownWordText(
  surface: string,
  headword: string,
  matchMode: NPlusOneMatchMode,
): string {
  return matchMode === 'surface' ? surface : headword;
}

function normalizePos1Tag(pos1: string | undefined): string {
  return typeof pos1 === 'string' ? pos1.trim() : '';
}

function splitNormalizedTagParts(normalizedTag: string): string[] {
  if (!normalizedTag) {
    return [];
  }

  return normalizedTag
    .split('|')
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
}

function isExcludedByTagSet(normalizedTag: string, exclusions: ReadonlySet<string>): boolean {
  const parts = splitNormalizedTagParts(normalizedTag);
  if (parts.length === 0) {
    return false;
  }

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

function isExcludedComponent(
  pos1: string | undefined,
  pos2: string | undefined,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): boolean {
  return (
    (typeof pos1 === 'string' && pos1Exclusions.has(pos1)) ||
    (typeof pos2 === 'string' && pos2Exclusions.has(pos2))
  );
}

function shouldAllowContentLedMergedTokenFrequency(
  normalizedPos1: string,
  normalizedPos2: string,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): boolean {
  const pos1Parts = splitNormalizedTagParts(normalizedPos1);
  if (pos1Parts.length < 2) {
    return false;
  }

  const pos2Parts = splitNormalizedTagParts(normalizedPos2);
  if (isExcludedComponent(pos1Parts[0], pos2Parts[0], pos1Exclusions, pos2Exclusions)) {
    return false;
  }

  const componentCount = Math.max(pos1Parts.length, pos2Parts.length);
  for (let index = 1; index < componentCount; index += 1) {
    if (!isExcludedComponent(pos1Parts[index], pos2Parts[index], pos1Exclusions, pos2Exclusions)) {
      return false;
    }
  }

  return true;
}

function shouldAllowOrdinalPrefixNounFrequency(token: MergedToken): boolean {
  const normalizedSurface = token.surface.trim();
  const normalizedHeadword = token.headword.trim();
  if (!normalizedSurface.startsWith('第') && !normalizedHeadword.startsWith('第')) {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePos1Tag(token.pos1));
  const pos2Parts = splitNormalizedTagParts(normalizePos2Tag(token.pos2));
  return (
    pos1Parts.length >= 2 &&
    pos1Parts[0] === '接頭詞' &&
    pos1Parts.slice(1).some((part) => part === '名詞') &&
    pos2Parts[0] === '数接続' &&
    pos2Parts.slice(1).some((part) => part === '数')
  );
}

function shouldAllowHonorificPrefixNounFrequency(token: MergedToken): boolean {
  const normalizedSurface = token.surface.trim();
  const normalizedHeadword = token.headword.trim();
  if (
    !['お', 'ご', '御'].some(
      (prefix) => normalizedSurface.startsWith(prefix) || normalizedHeadword.startsWith(prefix),
    )
  ) {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePos1Tag(token.pos1));
  const pos2Parts = splitNormalizedTagParts(normalizePos2Tag(token.pos2));
  return (
    pos1Parts.length >= 2 &&
    pos1Parts[0] === '接頭詞' &&
    pos1Parts.slice(1).some((part) => part === '名詞') &&
    pos2Parts[0] === '名詞接続'
  );
}

function shouldAllowDeterminerLedNounFrequency(
  normalizedPos1: string,
  normalizedPos2: string,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): boolean {
  const pos1Parts = splitNormalizedTagParts(normalizedPos1);
  if (pos1Parts.length < 2 || pos1Parts[0] !== '連体詞') {
    return false;
  }

  const pos2Parts = splitNormalizedTagParts(normalizedPos2);
  if (!isExcludedComponent(pos1Parts[0], pos2Parts[0], pos1Exclusions, pos2Exclusions)) {
    return false;
  }

  const componentCount = Math.max(pos1Parts.length, pos2Parts.length);
  for (let index = 1; index < componentCount; index += 1) {
    if (
      pos1Parts[index] === '名詞' &&
      !isExcludedComponent(pos1Parts[index], pos2Parts[index], pos1Exclusions, pos2Exclusions)
    ) {
      return true;
    }
  }

  return false;
}

function isFrequencyExcludedByPos(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): boolean {
  if (isSingleKanaFrequencyNoiseToken(token.surface)) {
    return true;
  }

  if (isKanaOnlyMixedFunctionContentToken(token, pos1Exclusions)) {
    return true;
  }

  const normalizedPos1 = normalizePos1Tag(token.pos1);
  const hasPos1 = normalizedPos1.length > 0;
  const normalizedPos2 = normalizePos2Tag(token.pos2);
  const hasPos2 = normalizedPos2.length > 0;
  const allowContentLedMergedToken = shouldAllowContentLedMergedTokenFrequency(
    normalizedPos1,
    normalizedPos2,
    pos1Exclusions,
    pos2Exclusions,
  );
  const allowDeterminerLedNounToken = shouldAllowDeterminerLedNounFrequency(
    normalizedPos1,
    normalizedPos2,
    pos1Exclusions,
    pos2Exclusions,
  );
  const allowOrdinalPrefixNounToken = shouldAllowOrdinalPrefixNounFrequency(token);
  const allowHonorificPrefixNounToken = shouldAllowHonorificPrefixNounFrequency(token);

  if (
    isExcludedByTagSet(normalizedPos1, pos1Exclusions) &&
    !allowContentLedMergedToken &&
    !allowDeterminerLedNounToken &&
    !allowOrdinalPrefixNounToken &&
    !allowHonorificPrefixNounToken
  ) {
    return true;
  }

  if (
    isExcludedByTagSet(normalizedPos2, pos2Exclusions) &&
    !allowContentLedMergedToken &&
    !allowDeterminerLedNounToken &&
    !allowOrdinalPrefixNounToken &&
    !allowHonorificPrefixNounToken
  ) {
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

export function shouldExcludeTokenFromVocabularyPersistence(
  token: MergedToken,
  options: Pick<AnnotationStageOptions, 'pos1Exclusions' | 'pos2Exclusions'> = {},
): boolean {
  const pos1Exclusions = resolvePos1Exclusions(options);
  const pos2Exclusions = resolvePos2Exclusions(options);

  return (
    sharedShouldExcludeTokenFromSubtitleAnnotations(token, { pos1Exclusions, pos2Exclusions }) ||
    (isFrequencyExcludedByPos(token, pos1Exclusions, pos2Exclusions) &&
      !isKanjiNonIndependentNounToken(token, pos1Exclusions))
  );
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
    code === 0x30fc ||
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

function isReduplicatedKanaSfxWithOptionalTrailingTo(text: string): boolean {
  const normalized = normalizeJlptTextForExclusion(text);
  if (!normalized) {
    return false;
  }

  if (isReduplicatedKanaSfx(normalized)) {
    return true;
  }

  if (normalized.length <= 1 || !normalized.endsWith('と')) {
    return false;
  }

  return isReduplicatedKanaSfx(normalized.slice(0, -1));
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

function isSingleKanaFrequencyNoiseToken(text: string | undefined): boolean {
  if (typeof text !== 'string') {
    return false;
  }

  const normalized = text.trim();
  if (!normalized) {
    return false;
  }

  const chars = [...normalized];
  return chars.length === 1 && isKanaChar(chars[0]!);
}

function isKanaOnlyText(text: string | undefined): boolean {
  if (typeof text !== 'string') {
    return false;
  }

  const normalized = text.trim();
  if (!normalized) {
    return false;
  }

  return [...normalized].every(isKanaChar);
}

function isKanaOnlyMixedFunctionContentToken(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
): boolean {
  if (!isKanaOnlyText(token.surface)) {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePos1Tag(token.pos1));
  return (
    pos1Parts.length >= 2 &&
    pos1Parts.some((part) => pos1Exclusions.has(part)) &&
    pos1Parts.some((part) => !pos1Exclusions.has(part))
  );
}

function isJlptEligibleToken(token: MergedToken): boolean {
  if (token.pos1 && shouldIgnoreJlptForMecabPos1(token.pos1)) {
    return false;
  }

  const candidates = [resolveJlptLookupText(token), token.surface, token.headword].filter(
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

export function shouldExcludeTokenFromSubtitleAnnotations(token: MergedToken): boolean {
  return sharedShouldExcludeTokenFromSubtitleAnnotations(token);
}

export function stripSubtitleAnnotationMetadata(
  token: MergedToken,
  options: AnnotationStageOptions = {},
): MergedToken {
  return sharedStripSubtitleAnnotationMetadata(token, options);
}

// Furigana-derived readings can be partial (kanji readings only, e.g. まあ for
// 待ち合わせてる); matching those against known words produces false positives,
// so the reading fallback requires a reading that plausibly covers the surface:
// at least as many characters as the surface, with the surface's kana appearing
// in order within the reading.
function isCompleteReadingForSurface(surface: string, reading: string): boolean {
  const surfaceChars = [...normalizeJlptTextForExclusion(surface)];
  const readingChars = [...normalizeJlptTextForExclusion(reading)];
  if (readingChars.length < surfaceChars.length) {
    return false;
  }

  let cursor = 0;
  for (const char of surfaceChars) {
    if (!isKanaChar(char)) {
      continue;
    }
    const foundAt = readingChars.indexOf(char, cursor);
    if (foundAt === -1) {
      return false;
    }
    cursor = foundAt + 1;
  }
  return true;
}

// Returns the token's trimmed reading only when it plausibly covers the surface
// (see isCompleteReadingForSurface); undefined otherwise. Shared so the
// known-word reading disambiguation and the reading fallback stay in sync if the
// validity rule changes.
function resolveCompleteTokenReading(token: MergedToken): string | undefined {
  const normalizedReading = token.reading.trim();
  if (!normalizedReading || !isCompleteReadingForSurface(token.surface, normalizedReading)) {
    return undefined;
  }
  return normalizedReading;
}

// Reading to disambiguate the known-word text match, or undefined when the
// token has no reading that describes the match text: in headword mode an
// inflected surface's reading does not match the dictionary form's reading,
// and partial furigana readings (see isCompleteReadingForSurface) would cause
// false negatives. Undefined falls back to text-only matching (fail-open).
function resolveKnownWordReadingForMatch(
  token: MergedToken,
  knownWordMatchMode: NPlusOneMatchMode,
): string | undefined {
  if (knownWordMatchMode === 'headword') {
    const headwordReading = token.headwordReading?.trim();
    if (headwordReading) {
      return headwordReading;
    }
    if (token.surface !== token.headword) {
      return undefined;
    }
  }

  return resolveCompleteTokenReading(token);
}

function computeTokenKnownStatus(
  token: MergedToken,
  isKnownWord: (text: string, reading?: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): boolean {
  const matchText = resolveKnownWordText(token.surface, token.headword, knownWordMatchMode);
  const matchReading = resolveKnownWordReadingForMatch(token, knownWordMatchMode);
  if (token.isKnown || (matchText ? isKnownWord(matchText, matchReading) : false)) {
    return true;
  }

  const fallbackReading = resolveCompleteTokenReading(token);
  if (!fallbackReading) {
    return false;
  }

  return fallbackReading !== matchText.trim() && isKnownWord(fallbackReading);
}

function filterTokenFrequencyRank(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): number | undefined {
  if (
    isFrequencyExcludedByPos(token, pos1Exclusions, pos2Exclusions) &&
    !isKanjiNonIndependentNounToken(token, pos1Exclusions)
  ) {
    return undefined;
  }

  if (typeof token.frequencyRank === 'number' && Number.isFinite(token.frequencyRank)) {
    return Math.max(1, Math.floor(token.frequencyRank));
  }

  return undefined;
}

function computeTokenJlptLevel(
  token: MergedToken,
  getJlptLevel: (text: string) => JlptLevel | null,
): JlptLevel | undefined {
  if (!isJlptEligibleToken(token)) {
    return undefined;
  }

  const primaryLevel = getCachedJlptLevel(resolveJlptLookupText(token), getJlptLevel);
  const fallbackLevel =
    primaryLevel === null ? getCachedJlptLevel(token.surface, getJlptLevel) : null;

  const level = primaryLevel ?? fallbackLevel ?? token.jlptLevel;
  return level ?? undefined;
}

function hasPrioritizedNameMatch(
  token: MergedToken,
  options: Pick<AnnotationStageOptions, 'nameMatchEnabled'>,
): boolean {
  return options.nameMatchEnabled !== false && token.isNameMatch === true;
}

export function annotateTokens(
  tokens: MergedToken[],
  deps: AnnotationStageDeps,
  options: AnnotationStageOptions = {},
): MergedToken[] {
  const pos1Exclusions = resolvePos1Exclusions(options);
  const pos2Exclusions = resolvePos2Exclusions(options);
  const knownWordsEnabled = options.knownWordsEnabled !== false;
  const nPlusOneEnabled = options.nPlusOneEnabled !== false;
  const nameMatchEnabled = options.nameMatchEnabled !== false;
  const frequencyEnabled = options.frequencyEnabled !== false;
  const jlptEnabled = options.jlptEnabled !== false;
  const shouldComputeKnownStatus = knownWordsEnabled || nPlusOneEnabled;
  const nPlusOneKnownStatuses: boolean[] = [];

  // Single pass: compute known word status, frequency filtering, and JLPT level together
  const annotated = tokens.map((token, index) => {
    const isKnownForMatching = shouldComputeKnownStatus
      ? computeTokenKnownStatus(token, deps.isKnownWord, deps.knownWordMatchMode)
      : false;
    nPlusOneKnownStatuses[index] = isKnownForMatching;

    const prioritizedNameMatch = nameMatchEnabled && token.isNameMatch === true;

    // A confirmed character-name match must survive the POS noise filter:
    // MeCab can tag a name like 平 as a prefix (接頭詞) depending on context,
    // which would otherwise strip the name match and its portrait.
    if (
      !prioritizedNameMatch &&
      sharedShouldExcludeTokenFromSubtitleAnnotations(token, {
        pos1Exclusions,
        pos2Exclusions,
      })
    ) {
      const strippedToken = sharedStripSubtitleAnnotationMetadata(token, {
        pos1Exclusions,
        pos2Exclusions,
      });
      return {
        ...strippedToken,
        isKnown: knownWordsEnabled ? isKnownForMatching : false,
      };
    }

    const frequencyRank =
      frequencyEnabled && !prioritizedNameMatch
        ? filterTokenFrequencyRank(token, pos1Exclusions, pos2Exclusions)
        : undefined;

    const jlptLevel =
      jlptEnabled && !prioritizedNameMatch
        ? computeTokenJlptLevel(token, deps.getJlptLevel)
        : undefined;

    return {
      ...token,
      isKnown: knownWordsEnabled ? isKnownForMatching : false,
      isNPlusOneTarget: nPlusOneEnabled && !prioritizedNameMatch ? token.isNPlusOneTarget : false,
      frequencyRank,
      jlptLevel,
    };
  });

  if (!nPlusOneEnabled) {
    return annotated;
  }

  const minSentenceWordsForNPlusOne = options.minSentenceWordsForNPlusOne;
  const sanitizedMinSentenceWordsForNPlusOne =
    minSentenceWordsForNPlusOne !== undefined &&
    Number.isInteger(minSentenceWordsForNPlusOne) &&
    minSentenceWordsForNPlusOne > 0
      ? minSentenceWordsForNPlusOne
      : 3;

  const nPlusOneMarked = nPlusOneEnabled
    ? markNPlusOneTargets(
        annotated.map((token, index) => ({
          ...token,
          isKnown: nPlusOneKnownStatuses[index] ?? false,
        })),
        sanitizedMinSentenceWordsForNPlusOne,
        pos1Exclusions,
        pos2Exclusions,
        options.sourceText,
      ).map((token, index) => ({
        ...annotated[index]!,
        isNPlusOneTarget: token.isNPlusOneTarget,
      }))
    : annotated;

  if (!nameMatchEnabled) {
    return nPlusOneMarked;
  }

  return nPlusOneMarked.map((token) =>
    hasPrioritizedNameMatch(token, options)
      ? {
          ...token,
          isNPlusOneTarget: false,
          frequencyRank: undefined,
          jlptLevel: undefined,
        }
      : token,
  );
}
