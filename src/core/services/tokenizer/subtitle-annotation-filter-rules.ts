import {
  DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG,
  resolveAnnotationPos1ExclusionSet,
} from '../../../token-pos1-exclusions';
import {
  DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG,
  resolveAnnotationPos2ExclusionSet,
} from '../../../token-pos2-exclusions';
import { MergedToken } from '../../../types';
import { shouldIgnoreJlptByTerm } from '../jlpt-token-filter';
import { isSubtitleGrammarEndingText } from './grammar-ending';
import {
  isAuxiliaryOnlyHelperSpan,
  isAuxiliaryStemGrammarTailToken,
  isExcludedTrailingParticleMergedToken,
  isKanaOnlyNonIndependentNounHelperMerge,
  isReduplicatedKanaSfxWithOptionalTrailingTo,
  isSingleKanaSurfaceFragment,
  isStandaloneAuxiliaryInflectionFragment,
  isStandaloneGrammarParticle,
  isStandaloneSuruTeGrammarHelper,
  isTrailingSmallTsuKanaSfx,
} from './subtitle-annotation-filter-support';
import {
  isContentTokenByPos,
  isPosTagExcluded,
  isTokenPos2Excluded,
  normalizeKana,
  normalizePosTag,
  splitPosTag,
} from './token-classification';

export interface SubtitleAnnotationFilterOptions {
  pos1Exclusions?: ReadonlySet<string>;
  pos2Exclusions?: ReadonlySet<string>;
}

export type SubtitleAnnotationRuleDecision = 'exclude' | 'keep' | 'pass';

export interface SubtitleAnnotationRuleContext {
  token: MergedToken;
  pos1Exclusions: ReadonlySet<string>;
  pos2Exclusions: ReadonlySet<string>;
  normalizedPos1: string;
  normalizedPos2: string;
  hasPos1: boolean;
  hasPos2: boolean;
}

type SubtitleAnnotationRuleData = Readonly<Record<string, readonly string[]>>;

export interface SubtitleAnnotationRule {
  readonly id: string;
  readonly description: string;
  readonly issueRef: string;
  readonly data: SubtitleAnnotationRuleData;
  test(context: SubtitleAnnotationRuleContext): SubtitleAnnotationRuleDecision;
}

function defineRule<TData extends SubtitleAnnotationRuleData>(definition: {
  id: string;
  description: string;
  issueRef: string;
  data: TData;
  test: (context: SubtitleAnnotationRuleContext, data: TData) => SubtitleAnnotationRuleDecision;
}): SubtitleAnnotationRule {
  const data = Object.freeze(
    Object.fromEntries(
      Object.entries(definition.data).map(([key, values]) => [key, Object.freeze([...values])]),
    ),
  ) as TData;
  return Object.freeze({
    id: definition.id,
    description: definition.description,
    issueRef: definition.issueRef,
    data,
    test: (context: SubtitleAnnotationRuleContext) => definition.test(context, data),
  });
}

export function createSubtitleAnnotationRuleContext(
  token: MergedToken,
  options: SubtitleAnnotationFilterOptions = {},
): SubtitleAnnotationRuleContext {
  const pos1Exclusions =
    options.pos1Exclusions ??
    resolveAnnotationPos1ExclusionSet(DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG);
  const pos2Exclusions =
    options.pos2Exclusions ??
    resolveAnnotationPos2ExclusionSet(DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG);
  const normalizedPos1 = normalizePosTag(token.pos1);
  const normalizedPos2 = normalizePosTag(token.pos2);
  return {
    token,
    pos1Exclusions,
    pos2Exclusions,
    normalizedPos1,
    normalizedPos2,
    hasPos1: normalizedPos1.length > 0,
    hasPos2: normalizedPos2.length > 0,
  };
}

function matchesExcludedTermOrPattern(token: MergedToken, terms: ReadonlySet<string>): boolean {
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
      terms.has(trimmed) ||
      terms.has(normalized) ||
      isSubtitleGrammarEndingText(trimmed) ||
      isSubtitleGrammarEndingText(normalized) ||
      shouldIgnoreJlptByTerm(trimmed) ||
      shouldIgnoreJlptByTerm(normalized) ||
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
  '何か',
  '何だ',
  '何も',
  '如何した',
  '有る',
  '在る',
  '様',
  '誰も',
  '貴方',
  'もんか',
  'ものか',
]);

const excludedTermRule = defineRule({
  id: 'excluded-term-or-pattern',
  description:
    'Exclude legacy subtitle stop terms, grammar endings, JLPT stop terms, and kana sound effects.',
  issueRef: '#19, #33, #57',
  data: {},
  test: ({ token }) =>
    matchesExcludedTermOrPattern(token, SUBTITLE_ANNOTATION_EXCLUDED_TERMS) ? 'exclude' : 'pass',
});

// Ordered, first-match-wins. issueRef records the introducing/fixing PR; early
// rules are legacy behavior inherited from the original #19 filter.
export const SUBTITLE_ANNOTATION_RULES: readonly SubtitleAnnotationRule[] = Object.freeze([
  defineRule({
    id: 'unparsed-run',
    description: 'Exclude hoverable parser gaps that have no Yomitan dictionary entry.',
    issueRef: '#153',
    data: {},
    test: ({ token }) => (token.isUnparsedRun === true ? 'exclude' : 'pass'),
  }),
  defineRule({
    id: 'configured-pos1-exclusion',
    description: 'Apply the configured primary part-of-speech exclusions.',
    issueRef: '#19',
    data: {},
    test: ({ token, pos1Exclusions }) =>
      isPosTagExcluded(token.pos1, pos1Exclusions) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'configured-pos2-exclusion',
    description: 'Apply configured secondary POS exclusions, preserving #150 kanji nouns.',
    issueRef: '#150',
    data: {},
    test: ({ token, pos1Exclusions, pos2Exclusions }) =>
      isTokenPos2Excluded(token, pos1Exclusions, pos2Exclusions) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'coarse-grammar-pos-fallback',
    description: 'Exclude coarse grammar POS when detailed MeCab tags are unavailable.',
    issueRef: '#19',
    data: {},
    test: ({ token, hasPos1, hasPos2, pos1Exclusions, pos2Exclusions }) =>
      !hasPos1 && !hasPos2 && !isContentTokenByPos(token, pos1Exclusions, pos2Exclusions)
        ? 'exclude'
        : 'pass',
  }),
  defineRule({
    id: 'auxiliary-stem-grammar-tail',
    description: 'Exclude merged grammar tails containing an auxiliary stem.',
    issueRef: '#19',
    data: { allowedPos1: ['名詞', '助動詞', '助詞'] },
    test: ({ token }, { allowedPos1 }) =>
      isAuxiliaryStemGrammarTailToken(token, allowedPos1) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'kana-non-independent-noun-helper',
    description: 'Exclude kana non-independent nouns merged with grammar helpers.',
    issueRef: '#56',
    data: { tailPos1: ['助詞', '助動詞'] },
    test: ({ token }, { tailPos1 }) =>
      isKanaOnlyNonIndependentNounHelperMerge(token, tailPos1) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'standalone-auxiliary-inflection',
    description: 'Exclude standalone kana auxiliary inflection fragments.',
    issueRef: '#57',
    data: { trailingPos1: ['助動詞'] },
    test: ({ token }, { trailingPos1 }) =>
      isStandaloneAuxiliaryInflectionFragment(token, trailingPos1) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'auxiliary-only-helper-span',
    description: 'Exclude kana helper spans without an independent lexical verb.',
    issueRef: '#57',
    data: { allowedPos1: ['助詞', '助動詞', '動詞'], lexicalVerbPos2: ['自立'] },
    test: ({ token }, { allowedPos1, lexicalVerbPos2 }) =>
      isAuxiliaryOnlyHelperSpan(token, allowedPos1, lexicalVerbPos2) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'standalone-suru-te-helper',
    description: 'Exclude standalone して grammar-helper fragments.',
    issueRef: '#57',
    data: {},
    test: ({ token }) => (isStandaloneSuruTeGrammarHelper(token) ? 'exclude' : 'pass'),
  }),
  defineRule({
    id: 'standalone-grammar-particle',
    description: 'Exclude standalone particle surfaces and connective particle phrases.',
    issueRef: '#57',
    data: {
      surfaces: [
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
      ],
      phrases: ['たって', 'だって'],
    },
    test: ({ token }, { surfaces, phrases }) =>
      isStandaloneGrammarParticle(token, surfaces, phrases) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'single-kana-fragment',
    description: 'Exclude isolated one-kana parser fragments.',
    issueRef: '#57',
    data: {},
    test: ({ token }) => (isSingleKanaSurfaceFragment(token) ? 'exclude' : 'pass'),
  }),
  defineRule({
    id: 'merged-trailing-quote-particle',
    description: 'Exclude lexical tokens merged only with a trailing quote-particle suffix.',
    issueRef: '#19',
    data: { suffixes: ['って', 'ってよ', 'ってね', 'ってな', 'ってさ', 'ってか', 'ってば'] },
    test: ({ token, pos1Exclusions }, { suffixes }) =>
      isExcludedTrailingParticleMergedToken(token, suffixes, pos1Exclusions) ? 'exclude' : 'pass',
  }),
  defineRule({
    id: 'lexical-kureru-keep',
    description: 'Keep lexical くれる before the legacy bare-くれ stop-term rule.',
    issueRef: '#57',
    data: {
      surfaces: ['くれ'],
      headwords: ['くれる'],
      pos1: ['動詞'],
      pos2: ['自立'],
    },
    test: ({ token }, data) => {
      const pos1 = splitPosTag(token.pos1);
      const pos2 = splitPosTag(token.pos2);
      return data.surfaces.includes(normalizeKana(token.surface)) &&
        data.headwords.includes(normalizeKana(token.headword)) &&
        pos1.length === 1 &&
        data.pos1.includes(pos1[0] ?? '') &&
        pos2.length === 1 &&
        data.pos2.includes(pos2[0] ?? '')
        ? 'keep'
        : 'pass';
    },
  }),
  excludedTermRule,
]);

export function evaluateSubtitleAnnotationRules(
  context: SubtitleAnnotationRuleContext,
): SubtitleAnnotationRuleDecision {
  for (const rule of SUBTITLE_ANNOTATION_RULES) {
    const decision = rule.test(context);
    if (decision !== 'pass') {
      return decision;
    }
  }
  return 'pass';
}
