import assert from 'node:assert/strict';
import test from 'node:test';
import { MergedToken, PartOfSpeech } from '../../../types';
import {
  createSubtitleAnnotationRuleContext,
  shouldExcludeTokenFromSubtitleAnnotations,
  SUBTITLE_ANNOTATION_RULES,
} from './subtitle-annotation-filter';

function makeToken(overrides: Partial<MergedToken> = {}): MergedToken {
  return {
    surface: '猫',
    reading: 'ネコ',
    headword: '猫',
    startPos: 0,
    endPos: 1,
    partOfSpeech: PartOfSpeech.noun,
    isMerged: false,
    isKnown: false,
    isNPlusOneTarget: false,
    ...overrides,
  };
}

test('configured POS rules consume the context exclusion sets in table order', () => {
  const token = makeToken({ pos1: '名詞', pos2: '一般' });
  const context = createSubtitleAnnotationRuleContext(token, {
    pos1Exclusions: new Set(['名詞']),
    pos2Exclusions: new Set(['一般']),
  });
  const pos1Rule = SUBTITLE_ANNOTATION_RULES.find(({ id }) => id === 'configured-pos1-exclusion');
  const pos2Rule = SUBTITLE_ANNOTATION_RULES.find(({ id }) => id === 'configured-pos2-exclusion');

  assert.equal(pos1Rule?.test(context), 'exclude');
  assert.equal(pos2Rule?.test(context), 'exclude');
  assert.equal(
    shouldExcludeTokenFromSubtitleAnnotations(token, {
      pos1Exclusions: context.pos1Exclusions,
      pos2Exclusions: context.pos2Exclusions,
    }),
    true,
  );
});

test('configured POS2 rule preserves the kanji non-independent noun exception', () => {
  const token = makeToken({ surface: '以外', headword: '以外', pos1: '名詞', pos2: '非自立' });
  const context = createSubtitleAnnotationRuleContext(token, {
    pos1Exclusions: new Set(),
    pos2Exclusions: new Set(['非自立']),
  });
  const pos2Rule = SUBTITLE_ANNOTATION_RULES.find(({ id }) => id === 'configured-pos2-exclusion');

  assert.equal(pos2Rule?.test(context), 'pass');
  assert.equal(
    shouldExcludeTokenFromSubtitleAnnotations(token, {
      pos1Exclusions: context.pos1Exclusions,
      pos2Exclusions: context.pos2Exclusions,
    }),
    false,
  );
});

test('trailing quote-particle rule honors configured POS1 exclusions', () => {
  const token = makeToken({ surface: '猫って', headword: '猫', pos1: '名詞|助詞' });
  const context = createSubtitleAnnotationRuleContext(token, {
    pos1Exclusions: new Set(['名詞']),
  });
  const trailingParticleRule = SUBTITLE_ANNOTATION_RULES.find(
    ({ id }) => id === 'merged-trailing-quote-particle',
  );

  assert.equal(trailingParticleRule?.test(context), 'pass');
});

test('configured POS2 rule preserves supplementary-plane kanji nouns', () => {
  const token = makeToken({ surface: '𠮟', headword: '𠮟', pos1: '名詞', pos2: '非自立' });

  assert.equal(
    shouldExcludeTokenFromSubtitleAnnotations(token, {
      pos1Exclusions: new Set(),
      pos2Exclusions: new Set(['非自立']),
    }),
    false,
  );
});
