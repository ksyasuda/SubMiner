import assert from 'node:assert/strict';
import test from 'node:test';
import { MergedToken, PartOfSpeech } from '../../../types';
import {
  createSubtitleAnnotationRuleContext,
  shouldExcludeTokenFromSubtitleAnnotations,
  SUBTITLE_ANNOTATION_RULES,
} from './subtitle-annotation-filter';
import { isKanaChar } from './token-classification';

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

test('subtitle annotation rules expose stable ordered provenance', () => {
  assert.deepEqual(
    SUBTITLE_ANNOTATION_RULES.map(({ id, issueRef }) => ({ id, issueRef })),
    [
      { id: 'unparsed-run', issueRef: '#153' },
      { id: 'configured-pos1-exclusion', issueRef: '#19' },
      { id: 'configured-pos2-exclusion', issueRef: '#150' },
      { id: 'coarse-grammar-pos-fallback', issueRef: '#19' },
      { id: 'auxiliary-stem-grammar-tail', issueRef: '#19' },
      { id: 'kana-non-independent-noun-helper', issueRef: '#56' },
      { id: 'standalone-auxiliary-inflection', issueRef: '#57' },
      { id: 'auxiliary-only-helper-span', issueRef: '#57' },
      { id: 'standalone-suru-te-helper', issueRef: '#57' },
      { id: 'standalone-grammar-particle', issueRef: '#57' },
      { id: 'single-kana-fragment', issueRef: '#57' },
      { id: 'merged-trailing-quote-particle', issueRef: '#19' },
      { id: 'lexical-kureru-keep', issueRef: '#57' },
      { id: 'excluded-term-or-pattern', issueRef: '#19, #33, #57' },
    ],
  );
  assert.equal(new Set(SUBTITLE_ANNOTATION_RULES.map(({ id }) => id)).size, 14);
  assert.ok(SUBTITLE_ANNOTATION_RULES.every(({ description }) => description.length > 0));
  assert.ok(
    SUBTITLE_ANNOTATION_RULES.every(
      ({ data }) => Object.isFrozen(data) && Object.values(data).every(Object.isFrozen),
    ),
  );
});

test('lexical kureru keep rule precedes and overrides the excluded-term rule', () => {
  const token = makeToken({
    surface: 'くれ',
    headword: 'くれる',
    reading: 'クレ',
    partOfSpeech: PartOfSpeech.verb,
    pos1: '動詞',
    pos2: '自立',
  });
  const context = createSubtitleAnnotationRuleContext(token);
  const keepIndex = SUBTITLE_ANNOTATION_RULES.findIndex(({ id }) => id === 'lexical-kureru-keep');
  const termIndex = SUBTITLE_ANNOTATION_RULES.findIndex(
    ({ id }) => id === 'excluded-term-or-pattern',
  );

  assert.ok(keepIndex >= 0 && keepIndex < termIndex);
  assert.ok(SUBTITLE_ANNOTATION_RULES[termIndex]?.data.terms?.includes('くれ'));
  assert.equal(SUBTITLE_ANNOTATION_RULES[keepIndex]?.test(context), 'keep');
  assert.equal(SUBTITLE_ANNOTATION_RULES[termIndex]?.test(context), 'exclude');
  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), false);
});

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

test('kana detection excludes katakana punctuation boundaries', () => {
  assert.equal(isKanaChar('゠'), false);
  assert.equal(isKanaChar('・'), false);
  assert.equal(isKanaChar('ァ'), true);
});
