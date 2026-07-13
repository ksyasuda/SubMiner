import { MergedToken } from '../../../types';
import {
  createSubtitleAnnotationRuleContext,
  evaluateSubtitleAnnotationRules,
} from './subtitle-annotation-filter-rules';
import type { SubtitleAnnotationFilterOptions } from './subtitle-annotation-filter-rules';

export {
  createSubtitleAnnotationRuleContext,
  SUBTITLE_ANNOTATION_EXCLUDED_TERMS,
  SUBTITLE_ANNOTATION_RULES,
} from './subtitle-annotation-filter-rules';
export type {
  SubtitleAnnotationFilterOptions,
  SubtitleAnnotationRule,
  SubtitleAnnotationRuleContext,
  SubtitleAnnotationRuleDecision,
} from './subtitle-annotation-filter-rules';
export { isKanjiNonIndependentNounToken } from './token-classification';

export function shouldExcludeTokenFromSubtitleAnnotations(
  token: MergedToken,
  options: SubtitleAnnotationFilterOptions = {},
): boolean {
  return (
    evaluateSubtitleAnnotationRules(createSubtitleAnnotationRuleContext(token, options)) ===
    'exclude'
  );
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
