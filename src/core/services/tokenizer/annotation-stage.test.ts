import assert from 'node:assert/strict';
import test from 'node:test';
import { MergedToken, PartOfSpeech } from '../../../types';
import { annotateTokens, AnnotationStageDeps } from './annotation-stage';

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

function makeDeps(overrides: Partial<AnnotationStageDeps> = {}): AnnotationStageDeps {
  return {
    isKnownWord: () => false,
    knownWordMatchMode: 'headword',
    getJlptLevel: () => null,
    ...overrides,
  };
}

test('annotateTokens known-word match mode uses headword vs surface', () => {
  const tokens = [makeToken({ surface: '食べた', headword: '食べる', reading: 'タベタ' })];
  const isKnownWord = (text: string): boolean => text === '食べる';

  const headwordResult = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord,
      knownWordMatchMode: 'headword',
    }),
  );
  const surfaceResult = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord,
      knownWordMatchMode: 'surface',
    }),
  );

  assert.equal(headwordResult[0]?.isKnown, true);
  assert.equal(surfaceResult[0]?.isKnown, false);
});

test('annotateTokens excludes frequency for particle/bound_auxiliary and pos1 exclusions', () => {
  const tokens = [
    makeToken({
      surface: 'は',
      headword: 'は',
      partOfSpeech: PartOfSpeech.particle,
      frequencyRank: 3,
    }),
    makeToken({
      surface: 'です',
      headword: 'です',
      partOfSpeech: PartOfSpeech.bound_auxiliary,
      startPos: 1,
      endPos: 3,
      frequencyRank: 4,
    }),
    makeToken({
      surface: 'の',
      headword: 'の',
      partOfSpeech: PartOfSpeech.other,
      pos1: '助詞',
      startPos: 3,
      endPos: 4,
      frequencyRank: 5,
    }),
    makeToken({
      surface: '猫',
      headword: '猫',
      partOfSpeech: PartOfSpeech.noun,
      startPos: 4,
      endPos: 5,
      frequencyRank: 11,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps());

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[1]?.frequencyRank, undefined);
  assert.equal(result[2]?.frequencyRank, undefined);
  assert.equal(result[3]?.frequencyRank, 11);
});

test('annotateTokens preserves existing frequency rank when frequency is enabled', () => {
  const tokens = [makeToken({ surface: '猫', headword: '猫', frequencyRank: 42 })];

  const result = annotateTokens(tokens, makeDeps());

  assert.equal(result[0]?.frequencyRank, 42);
});

test('annotateTokens drops invalid frequency rank values', () => {
  const tokens = [makeToken({ surface: '猫', headword: '猫', frequencyRank: Number.NaN })];
  const result = annotateTokens(tokens, makeDeps());
  assert.equal(result[0]?.frequencyRank, undefined);
});

test('annotateTokens clears frequency rank when frequency is disabled', () => {
  const tokens = [makeToken({ surface: '猫', headword: '猫', frequencyRank: 42 })];
  const result = annotateTokens(tokens, makeDeps(), { frequencyEnabled: false });
  assert.equal(result[0]?.frequencyRank, undefined);
});

test('annotateTokens handles JLPT disabled and eligibility exclusion paths', () => {
  let disabledLookupCalls = 0;
  const disabledResult = annotateTokens(
    [makeToken({ surface: '猫', headword: '猫' })],
    makeDeps({
      getJlptLevel: () => {
        disabledLookupCalls += 1;
        return 'N5';
      },
    }),
    { jlptEnabled: false },
  );
  assert.equal(disabledResult[0]?.jlptLevel, undefined);
  assert.equal(disabledLookupCalls, 0);

  let excludedLookupCalls = 0;
  const excludedResult = annotateTokens(
    [
      makeToken({
        surface: '！',
        headword: '！',
        reading: '',
        pos1: '記号',
        partOfSpeech: PartOfSpeech.symbol,
      }),
    ],
    makeDeps({
      getJlptLevel: () => {
        excludedLookupCalls += 1;
        return 'N5';
      },
    }),
  );
  assert.equal(excludedResult[0]?.jlptLevel, undefined);
  assert.equal(excludedLookupCalls, 0);
});

test('annotateTokens N+1 handoff marks expected target when threshold is satisfied', () => {
  const tokens = [
    makeToken({ surface: '私', headword: '私', startPos: 0, endPos: 1 }),
    makeToken({ surface: '猫', headword: '猫', startPos: 1, endPos: 2 }),
    makeToken({
      surface: '見る',
      headword: '見る',
      partOfSpeech: PartOfSpeech.verb,
      startPos: 2,
      endPos: 4,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '見る',
    }),
    { minSentenceWordsForNPlusOne: 3 },
  );

  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.isNPlusOneTarget, true);
  assert.equal(result[2]?.isNPlusOneTarget, false);
});

test('annotateTokens N+1 minimum sentence words counts only eligible word tokens', () => {
  const tokens = [
    makeToken({ surface: '猫', headword: '猫', startPos: 0, endPos: 1 }),
    makeToken({
      surface: 'が',
      headword: 'が',
      partOfSpeech: PartOfSpeech.particle,
      pos1: '助詞',
      startPos: 1,
      endPos: 2,
    }),
    makeToken({
      surface: 'です',
      headword: 'です',
      partOfSpeech: PartOfSpeech.bound_auxiliary,
      pos1: '助動詞',
      startPos: 2,
      endPos: 4,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'が' || text === 'です',
    }),
    { minSentenceWordsForNPlusOne: 3 },
  );

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[1]?.isKnown, true);
  assert.equal(result[2]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens applies configured pos1 exclusions to both frequency and N+1', () => {
  const tokens = [
    makeToken({
      surface: '猫',
      headword: '猫',
      pos1: '名詞',
      frequencyRank: 21,
      startPos: 0,
      endPos: 1,
    }),
    makeToken({
      surface: '走る',
      headword: '走る',
      pos1: '動詞',
      partOfSpeech: PartOfSpeech.verb,
      startPos: 1,
      endPos: 3,
      frequencyRank: 22,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '走る',
    }),
    {
      minSentenceWordsForNPlusOne: 1,
      pos1Exclusions: new Set(['名詞']),
    },
  );

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[1]?.frequencyRank, 22);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.isNPlusOneTarget, false);
});

test('annotateTokens allows previously default-excluded pos1 when removed from effective set', () => {
  const tokens = [
    makeToken({
      surface: 'まで',
      headword: 'まで',
      partOfSpeech: PartOfSpeech.other,
      pos1: '助詞',
      startPos: 0,
      endPos: 2,
      frequencyRank: 8,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
    pos1Exclusions: new Set(),
  });

  assert.equal(result[0]?.frequencyRank, 8);
  assert.equal(result[0]?.isNPlusOneTarget, true);
});

test('annotateTokens excludes default non-independent pos2 from frequency and N+1', () => {
  const tokens = [
    makeToken({
      surface: 'になれば',
      headword: 'なる',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '非自立',
      startPos: 0,
      endPos: 4,
      frequencyRank: 7,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens keeps frequency for kanji noun tokens even when mecab marks them non-independent', () => {
  const tokens = [
    makeToken({
      surface: '者',
      reading: 'もの',
      headword: '者',
      partOfSpeech: PartOfSpeech.other,
      pos1: '名詞',
      pos2: '非自立',
      pos3: '一般',
      startPos: 0,
      endPos: 1,
      frequencyRank: 475,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, 475);
});

test('annotateTokens excludes likely kana SFX tokens from frequency when POS tags are missing', () => {
  const tokens = [
    makeToken({
      surface: 'ぐわっ',
      reading: 'ぐわっ',
      headword: 'ぐわっ',
      pos1: '',
      pos2: '',
      frequencyRank: 12,
      startPos: 0,
      endPos: 3,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
});

test('annotateTokens excludes single hiragana and katakana tokens from frequency when POS tags are missing', () => {
  const tokens = [
    makeToken({
      surface: 'た',
      reading: 'た',
      headword: 'た',
      pos1: '',
      pos2: '',
      partOfSpeech: PartOfSpeech.other,
      frequencyRank: 21,
      startPos: 0,
      endPos: 1,
    }),
    makeToken({
      surface: 'ア',
      reading: 'ア',
      headword: 'ア',
      pos1: '',
      pos2: '',
      partOfSpeech: PartOfSpeech.other,
      frequencyRank: 22,
      startPos: 1,
      endPos: 2,
    }),
    makeToken({
      surface: '山',
      reading: 'やま',
      headword: '山',
      pos1: '',
      pos2: '',
      partOfSpeech: PartOfSpeech.other,
      frequencyRank: 23,
      startPos: 2,
      endPos: 3,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[1]?.frequencyRank, undefined);
  assert.equal(result[2]?.frequencyRank, 23);
});

test('annotateTokens keeps frequency when mecab tags classify token as content-bearing', () => {
  const tokens = [
    makeToken({
      surface: 'ふふ',
      headword: 'ふふ',
      pos1: '動詞',
      pos2: '自立',
      frequencyRank: 3014,
      startPos: 0,
      endPos: 2,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, 3014);
});

test('annotateTokens allows previously default-excluded pos2 when removed from effective set', () => {
  const tokens = [
    makeToken({
      surface: 'になれば',
      headword: 'なる',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '非自立',
      startPos: 0,
      endPos: 4,
      frequencyRank: 9,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
    pos2Exclusions: new Set(),
  });

  assert.equal(result[0]?.frequencyRank, 9);
  assert.equal(result[0]?.isNPlusOneTarget, true);
});

test('annotateTokens excludes composite function/content tokens from frequency but keeps N+1 eligible', () => {
  const tokens = [
    makeToken({
      surface: 'になれば',
      headword: 'なる',
      pos1: '助詞|動詞',
      pos2: '格助詞|自立|接続助詞',
      startPos: 0,
      endPos: 4,
      frequencyRank: 5,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.isNPlusOneTarget, true);
});

test('annotateTokens excludes composite tokens when all component pos tags are excluded', () => {
  const tokens = [
    makeToken({
      surface: 'けど',
      headword: 'けど',
      pos1: '助詞|助詞',
      pos2: '接続助詞|終助詞',
      startPos: 0,
      endPos: 2,
      frequencyRank: 6,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});
