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
  const lookupCalls: string[] = [];
  const tokens = [
    makeToken({ surface: 'は', headword: 'は', partOfSpeech: PartOfSpeech.particle }),
    makeToken({
      surface: 'です',
      headword: 'です',
      partOfSpeech: PartOfSpeech.bound_auxiliary,
      startPos: 1,
      endPos: 3,
    }),
    makeToken({
      surface: 'の',
      headword: 'の',
      partOfSpeech: PartOfSpeech.other,
      pos1: '助詞',
      startPos: 3,
      endPos: 4,
    }),
    makeToken({
      surface: '猫',
      headword: '猫',
      partOfSpeech: PartOfSpeech.noun,
      startPos: 4,
      endPos: 5,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      getFrequencyRank: (text) => {
        lookupCalls.push(text);
        return text === '猫' ? 11 : 999;
      },
    }),
  );

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[1]?.frequencyRank, undefined);
  assert.equal(result[2]?.frequencyRank, undefined);
  assert.equal(result[3]?.frequencyRank, 11);
  assert.deepEqual(lookupCalls, ['猫']);
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
