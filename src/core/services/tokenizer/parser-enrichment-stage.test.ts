import test from 'node:test';
import assert from 'node:assert/strict';
import { MergedToken, PartOfSpeech } from '../../../types';
import { enrichTokensWithMecabPos1 } from './parser-enrichment-stage';

function makeToken(overrides: Partial<MergedToken>): MergedToken {
  return {
    surface: 'token',
    reading: '',
    headword: 'token',
    startPos: 0,
    endPos: 1,
    partOfSpeech: PartOfSpeech.other,
    isMerged: true,
    isKnown: false,
    isNPlusOneTarget: false,
    pos1: '',
    ...overrides,
  };
}

test('enrichTokensWithMecabPos1 picks pos1 by best overlap when no surface match exists', () => {
  const tokens = [makeToken({ surface: 'grouped', startPos: 2, endPos: 7 })];
  const mecabTokens = [
    makeToken({ surface: 'left', startPos: 0, endPos: 4, pos1: 'A' }),
    makeToken({ surface: 'right', startPos: 2, endPos: 6, pos1: 'B' }),
  ];

  const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
  assert.equal(enriched[0]?.pos1, 'B');
});

test('enrichTokensWithMecabPos1 fills missing pos1 using surface-sequence fallback', () => {
  const tokens = [makeToken({ surface: ' は ', startPos: 10, endPos: 13 })];
  const mecabTokens = [makeToken({ surface: 'は', startPos: 0, endPos: 1, pos1: '助詞' })];

  const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
  assert.equal(enriched[0]?.pos1, '助詞');
});

test('enrichTokensWithMecabPos1 passes through unchanged when mecab tokens are null or empty', () => {
  const tokens = [makeToken({ surface: '猫', startPos: 0, endPos: 1 })];

  const nullResult = enrichTokensWithMecabPos1(tokens, null);
  assert.strictEqual(nullResult, tokens);

  const emptyResult = enrichTokensWithMecabPos1(tokens, []);
  assert.strictEqual(emptyResult, tokens);
});
