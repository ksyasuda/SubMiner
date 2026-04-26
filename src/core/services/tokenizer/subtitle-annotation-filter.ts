import {
  DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG,
  resolveAnnotationPos1ExclusionSet,
} from '../../../token-pos1-exclusions';
import {
  DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG,
  resolveAnnotationPos2ExclusionSet,
} from '../../../token-pos2-exclusions';
import { MergedToken, PartOfSpeech } from '../../../types';
import { shouldIgnoreJlptByTerm } from '../jlpt-token-filter';

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;

const SUBTITLE_ANNOTATION_EXCLUDED_TERMS = new Set([
  'あ',
  'ああ',
  'ええ',
  'うう',
  'おお',
  'はあ',
  'はは',
  'へえ',
  'ふう',
  'ほう',
]);
const SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_PREFIXES = ['ん', 'の', 'なん', 'なの'];
const SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_CORES = [
  'だ',
  'です',
  'でした',
  'だった',
  'では',
  'じゃ',
  'でしょう',
  'だろう',
] as const;
const SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_TRAILING_PARTICLES = [
  '',
  'か',
  'ね',
  'よ',
  'な',
  'けど',
  'よね',
  'かな',
  'かね',
] as const;
const SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_THOUGHT_SUFFIXES = [
  'か',
  'かな',
  'かね',
] as const;
const SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDINGS = new Set(
  SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_PREFIXES.flatMap((prefix) =>
    SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_CORES.flatMap((core) =>
      SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_TRAILING_PARTICLES.map(
        (particle) => `${prefix}${core}${particle}`,
      ),
    ),
  ),
);
const SUBTITLE_ANNOTATION_EXCLUDED_TRAILING_PARTICLE_SUFFIXES = new Set([
  'って',
  'ってよ',
  'ってね',
  'ってな',
  'ってさ',
  'ってか',
  'ってば',
]);
const AUXILIARY_STEM_GRAMMAR_TAIL_POS1 = new Set(['名詞', '助動詞', '助詞']);
const NON_INDEPENDENT_NOUN_HELPER_TAIL_POS1 = new Set(['助詞', '助動詞']);

export interface SubtitleAnnotationFilterOptions {
  pos1Exclusions?: ReadonlySet<string>;
  pos2Exclusions?: ReadonlySet<string>;
}

function normalizePosTag(pos: string | undefined): string {
  return typeof pos === 'string' ? pos.trim() : '';
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

function resolvePos1Exclusions(options: SubtitleAnnotationFilterOptions = {}): ReadonlySet<string> {
  if (options.pos1Exclusions) {
    return options.pos1Exclusions;
  }

  return resolveAnnotationPos1ExclusionSet(DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG);
}

function resolvePos2Exclusions(options: SubtitleAnnotationFilterOptions = {}): ReadonlySet<string> {
  if (options.pos2Exclusions) {
    return options.pos2Exclusions;
  }

  return resolveAnnotationPos2ExclusionSet(DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG);
}

function normalizeKana(text: string): string {
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

function isTrailingSmallTsuKanaSfx(text: string): boolean {
  const normalized = normalizeKana(text);
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
  const normalized = normalizeKana(text);
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
  const normalized = normalizeKana(text);
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

function isExcludedTrailingParticleMergedToken(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  const normalizedHeadword = normalizeKana(token.headword);
  if (
    !normalizedSurface ||
    !normalizedHeadword ||
    !normalizedSurface.startsWith(normalizedHeadword)
  ) {
    return false;
  }

  const suffix = normalizedSurface.slice(normalizedHeadword.length);
  if (!SUBTITLE_ANNOTATION_EXCLUDED_TRAILING_PARTICLE_SUFFIXES.has(suffix)) {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePosTag(token.pos1));
  if (pos1Parts.length < 2) {
    return false;
  }

  const [leadingPos1, ...trailingPos1] = pos1Parts;
  if (!leadingPos1 || resolvePos1Exclusions().has(leadingPos1)) {
    return false;
  }

  return trailingPos1.length > 0 && trailingPos1.every((part) => part === '助詞');
}

function isAuxiliaryStemGrammarTailToken(token: MergedToken): boolean {
  const pos1Parts = splitNormalizedTagParts(normalizePosTag(token.pos1));
  if (
    pos1Parts.length === 0 ||
    !pos1Parts.every((part) => AUXILIARY_STEM_GRAMMAR_TAIL_POS1.has(part))
  ) {
    return false;
  }

  const pos3Parts = splitNormalizedTagParts(normalizePosTag(token.pos3));
  return pos3Parts.includes('助動詞語幹');
}

function isKanaOnlyNonIndependentNounHelperMerge(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  const normalizedHeadword = normalizeKana(token.headword);
  if (
    !normalizedSurface ||
    !normalizedHeadword ||
    normalizedSurface === normalizedHeadword ||
    ![...normalizedSurface].every(isKanaChar)
  ) {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePosTag(token.pos1));
  if (pos1Parts.length < 2 || pos1Parts[0] !== '名詞') {
    return false;
  }

  const pos2Parts = splitNormalizedTagParts(normalizePosTag(token.pos2));
  if (pos2Parts[0] !== '非自立') {
    return false;
  }

  return pos1Parts.slice(1).every((part) => NON_INDEPENDENT_NOUN_HELPER_TAIL_POS1.has(part));
}

function isExcludedByTerm(token: MergedToken): boolean {
  const candidates = [token.surface, token.reading, token.headword].filter(
    (candidate): candidate is string => typeof candidate === 'string' && candidate.length > 0,
  );

  for (const candidate of candidates) {
    const trimmed = candidate.trim();
    if (!trimmed) {
      continue;
    }

    const normalized = normalizeKana(trimmed);
    if (!normalized) {
      continue;
    }

    if (
      SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_PREFIXES.some((prefix) =>
        SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDING_THOUGHT_SUFFIXES.some(
          (suffix) => normalized === `${prefix}${suffix}`,
        ),
      )
    ) {
      return true;
    }

    if (
      SUBTITLE_ANNOTATION_EXCLUDED_TERMS.has(trimmed) ||
      SUBTITLE_ANNOTATION_EXCLUDED_TERMS.has(normalized) ||
      SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDINGS.has(trimmed) ||
      SUBTITLE_ANNOTATION_EXCLUDED_EXPLANATORY_ENDINGS.has(normalized) ||
      shouldIgnoreJlptByTerm(trimmed) ||
      shouldIgnoreJlptByTerm(normalized)
    ) {
      return true;
    }

    if (
      isTrailingSmallTsuKanaSfx(trimmed) ||
      isTrailingSmallTsuKanaSfx(normalized) ||
      isReduplicatedKanaSfxWithOptionalTrailingTo(trimmed) ||
      isReduplicatedKanaSfxWithOptionalTrailingTo(normalized)
    ) {
      return true;
    }
  }

  return false;
}

export function shouldExcludeTokenFromSubtitleAnnotations(
  token: MergedToken,
  options: SubtitleAnnotationFilterOptions = {},
): boolean {
  const pos1Exclusions = resolvePos1Exclusions(options);
  const pos2Exclusions = resolvePos2Exclusions(options);
  const normalizedPos1 = normalizePosTag(token.pos1);
  const normalizedPos2 = normalizePosTag(token.pos2);
  const hasPos1 = normalizedPos1.length > 0;
  const hasPos2 = normalizedPos2.length > 0;

  if (isExcludedByTagSet(normalizedPos1, pos1Exclusions)) {
    return true;
  }

  if (isExcludedByTagSet(normalizedPos2, pos2Exclusions)) {
    return true;
  }

  if (
    !hasPos1 &&
    !hasPos2 &&
    (token.partOfSpeech === PartOfSpeech.particle ||
      token.partOfSpeech === PartOfSpeech.bound_auxiliary ||
      token.partOfSpeech === PartOfSpeech.symbol)
  ) {
    return true;
  }

  if (isAuxiliaryStemGrammarTailToken(token)) {
    return true;
  }

  if (isKanaOnlyNonIndependentNounHelperMerge(token)) {
    return true;
  }

  if (isExcludedTrailingParticleMergedToken(token)) {
    return true;
  }

  return isExcludedByTerm(token);
}

export function stripSubtitleAnnotationMetadata(
  token: MergedToken,
  options: SubtitleAnnotationFilterOptions = {},
): MergedToken {
  if (!shouldExcludeTokenFromSubtitleAnnotations(token, options)) {
    return token;
  }

  return {
    ...token,
    isKnown: false,
    isNPlusOneTarget: false,
    isNameMatch: false,
    jlptLevel: undefined,
    frequencyRank: undefined,
  };
}
