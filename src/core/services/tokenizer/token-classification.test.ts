import assert from 'node:assert/strict';
import test from 'node:test';
import { MergedToken, PartOfSpeech } from '../../../types';
import {
  isContentTokenByPos,
  isKanaCandidateIgnorableChar,
  isKanaCandidateText,
  isKanaChar,
  isKanaOnlyText,
  isTokenPos2Excluded,
  normalizeKana,
} from './token-classification';

const POS1_EXCLUSIONS = new Set(['助詞']);
const POS2_EXCLUSIONS = new Set(['非自立']);

function makeNoun(surface: string): MergedToken {
  return {
    surface,
    reading: surface,
    headword: surface,
    startPos: 0,
    endPos: surface.length,
    partOfSpeech: PartOfSpeech.noun,
    pos1: '名詞',
    pos2: '非自立',
    isMerged: false,
    isKnown: false,
    isNPlusOneTarget: false,
  };
}

test('kana normalization folds halfwidth kana, composing the voiced pairs', () => {
  // ｶ + ﾞ is two code points for one character: without composing them, a
  // halfwidth word counts as longer than the reading that spells it, which
  // disqualifies the reading from known-word matching.
  assert.equal(normalizeKana('ｶﾞｸ'), normalizeKana('ガク'));
  assert.equal(normalizeKana('ﾊﾟﾝ'), normalizeKana('パン'));
  assert.equal(normalizeKana('ﾐﾅﾄ'), 'みなと');
  assert.ok(isKanaOnlyText('ｶﾞｸ'));
});

test('kana normalization leaves characters other than halfwidth kana alone', () => {
  // The composition is scoped to the halfwidth runs: applied to the whole
  // string, NFKC would also rewrite these into something the dictionary, the
  // known-word list, and the frequency data were never keyed on.
  assert.equal(normalizeKana('①ｶﾞ'), '①が');
  assert.equal(normalizeKana('Ａｶﾞ'), 'Ａが');
  assert.equal(normalizeKana('㍑ｶﾞ'), '㍑が');
  assert.equal(normalizeKana('ﬁｶﾞ'), 'ﬁが');
});

test('kana classification excludes the katakana-hiragana double hyphen', () => {
  assert.equal(isKanaChar('゠'), false);
  assert.equal(isKanaOnlyText('゠'), false);
});

test('POS classification keeps kanji non-independent nouns as content', () => {
  const token = makeNoun('日');

  assert.equal(isTokenPos2Excluded(token, POS1_EXCLUSIONS, POS2_EXCLUSIONS), false);
  assert.equal(isContentTokenByPos(token, POS1_EXCLUSIONS, POS2_EXCLUSIONS), true);
});

test('POS classification excludes kana non-independent nouns', () => {
  const token = makeNoun('こと');

  assert.equal(isTokenPos2Excluded(token, POS1_EXCLUSIONS, POS2_EXCLUSIONS), true);
  assert.equal(isContentTokenByPos(token, POS1_EXCLUSIONS, POS2_EXCLUSIONS), false);
});

test('kana candidate classification allows punctuation around kana only', () => {
  assert.equal(isKanaCandidateIgnorableChar('！'), true);
  assert.equal(isKanaCandidateIgnorableChar('猫'), false);
  assert.equal(isKanaCandidateText('「かな！？」'), true);
  assert.equal(isKanaCandidateText('「！？」'), false);
  assert.equal(isKanaCandidateText('かな猫'), false);
});
