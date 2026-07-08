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
import { isSubtitleGrammarEndingText } from './grammar-ending';

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;

const STANDALONE_GRAMMAR_PARTICLE_PHRASES = ['たって', 'だって'] as const;
const STANDALONE_GRAMMAR_PARTICLE_PHRASES_SET: ReadonlySet<string> = new Set(
  STANDALONE_GRAMMAR_PARTICLE_PHRASES,
);

export const SUBTITLE_ANNOTATION_EXCLUDED_TERMS = new Set([
  'あ',
  'ああ',
  'ある',
  'あなた',
  'あんた',
  'ええ',
  'うう',
  'おお',
  'おい',
  'お前',
  'こいつ',
  'こっち',
  'くれ',
  'じゃない',
  'そうだ',
  'たち',
  'である',
  'どこか',
  'なんか',
  'べき',
  'って',
  'はあ',
  'はぁ',
  'はは',
  'へえ',
  'ふう',
  'ほう',
  'やはり',
  '何か',
  '何だ',
  '何も',
  '如何した',
  '有る',
  '在る',
  '様',
  '確かに',
  '誰も',
  '貴方',
  'もんか',
  'ものか',
]);
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
const AUXILIARY_INFLECTION_TRAILING_POS1 = new Set(['助動詞']);
const AUXILIARY_HELPER_SPAN_POS1 = new Set(['助詞', '助動詞', '動詞']);
const LEXICAL_VERB_POS2 = new Set(['自立']);
const STANDALONE_GRAMMAR_PARTICLE_SURFACES = new Set([
  'か',
  'が',
  'さ',
  'し',
  'ぞ',
  'ぜ',
  'と',
  'な',
  'に',
  'ね',
  'の',
  'は',
  'へ',
  'も',
  'や',
  'よ',
  'を',
]);
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

function isKanaOnlyText(text: string): boolean {
  const normalized = normalizeKana(text);
  return normalized.length > 0 && [...normalized].every(isKanaChar);
}

function isLexicalKureruVerb(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  const normalizedHeadword = normalizeKana(token.headword);
  const pos1Parts = splitNormalizedTagParts(normalizePosTag(token.pos1));
  const pos2Parts = splitNormalizedTagParts(normalizePosTag(token.pos2));
  return (
    normalizedSurface === 'くれ' &&
    normalizedHeadword === 'くれる' &&
    pos1Parts.length === 1 &&
    pos1Parts[0] === '動詞' &&
    pos2Parts.length === 1 &&
    pos2Parts[0] === '自立'
  );
}

function isStandaloneAuxiliaryInflectionFragment(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  if (!isKanaOnlyText(normalizedSurface)) {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePosTag(token.pos1));
  if (pos1Parts.length === 0) {
    return false;
  }

  if (pos1Parts.every((part) => part === '助動詞')) {
    return true;
  }

  const pos2Parts = splitNormalizedTagParts(normalizePosTag(token.pos2));
  return (
    pos1Parts[0] === '動詞' &&
    pos2Parts[0] === '接尾' &&
    pos1Parts.slice(1).every((part) => AUXILIARY_INFLECTION_TRAILING_POS1.has(part))
  );
}

function isAuxiliaryOnlyHelperSpan(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  const normalizedHeadword = normalizeKana(token.headword);
  if (!isKanaOnlyText(normalizedSurface) || !isKanaOnlyText(normalizedHeadword)) {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePosTag(token.pos1));
  if (
    pos1Parts.length === 0 ||
    !pos1Parts.every((part) => AUXILIARY_HELPER_SPAN_POS1.has(part)) ||
    !pos1Parts.includes('助詞') ||
    !pos1Parts.includes('動詞')
  ) {
    return false;
  }

  const pos2Parts = splitNormalizedTagParts(normalizePosTag(token.pos2));
  return !pos2Parts.some((part) => LEXICAL_VERB_POS2.has(part));
}

function isStandaloneSuruTeGrammarHelper(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  const normalizedHeadword = normalizeKana(token.headword);
  if (!normalizedSurface.startsWith('して') || normalizedHeadword !== 'する') {
    return false;
  }

  const pos1Parts = splitNormalizedTagParts(normalizePosTag(token.pos1));
  return (
    isKanaOnlyText(normalizedSurface) && (pos1Parts.length === 0 || pos1Parts.includes('動詞'))
  );
}

function isStandaloneGrammarParticle(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  const normalizedHeadword = normalizeKana(token.headword);
  return (
    normalizedSurface === normalizedHeadword &&
    (STANDALONE_GRAMMAR_PARTICLE_SURFACES.has(normalizedSurface) ||
      STANDALONE_GRAMMAR_PARTICLE_PHRASES_SET.has(normalizedSurface))
  );
}

function isSingleKanaSurfaceFragment(token: MergedToken): boolean {
  const normalizedSurface = normalizeKana(token.surface);
  const chars = [...normalizedSurface];
  return chars.length === 1 && chars.every(isKanaChar);
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
      SUBTITLE_ANNOTATION_EXCLUDED_TERMS.has(trimmed) ||
      SUBTITLE_ANNOTATION_EXCLUDED_TERMS.has(normalized) ||
      isSubtitleGrammarEndingText(trimmed) ||
      isSubtitleGrammarEndingText(normalized) ||
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

  if (isStandaloneAuxiliaryInflectionFragment(token)) {
    return true;
  }

  if (isAuxiliaryOnlyHelperSpan(token)) {
    return true;
  }

  if (isStandaloneSuruTeGrammarHelper(token)) {
    return true;
  }

  if (isStandaloneGrammarParticle(token)) {
    return true;
  }

  if (isSingleKanaSurfaceFragment(token)) {
    return true;
  }

  if (isExcludedTrailingParticleMergedToken(token)) {
    return true;
  }

  if (isLexicalKureruVerb(token)) {
    return false;
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

  const strippedToken = {
    ...token,
    isKnown: false,
    isNPlusOneTarget: false,
    isNameMatch: false,
    jlptLevel: undefined,
    frequencyRank: undefined,
  };

  if ('characterImage' in strippedToken) {
    strippedToken.characterImage = undefined;
  }

  return strippedToken;
}
