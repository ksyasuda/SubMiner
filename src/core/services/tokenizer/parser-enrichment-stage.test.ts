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
    makeToken({ surface: 'left', startPos: 0, endPos: 4, pos1: 'A', pos2: 'L2' }),
    makeToken({ surface: 'right', startPos: 2, endPos: 6, pos1: 'B', pos2: '非自立' }),
  ];

  const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
  assert.equal(enriched[0]?.pos1, 'A|B');
  assert.equal(enriched[0]?.pos2, 'L2|非自立');
});

test('enrichTokensWithMecabPos1 fills missing pos1 using surface-sequence fallback', () => {
  const tokens = [makeToken({ surface: ' は ', startPos: 10, endPos: 13 })];
  const mecabTokens = [makeToken({ surface: 'は', startPos: 0, endPos: 1, pos1: '助詞' })];

  const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
  assert.equal(enriched[0]?.pos1, '助詞');
});

test('enrichTokensWithMecabPos1 backfills blank pos2 and pos3 fields', () => {
  const tokens = [
    makeToken({
      surface: 'は',
      startPos: 0,
      endPos: 1,
      pos1: '助詞',
      pos2: '',
      pos3: ' ',
    }),
  ];
  const mecabTokens = [
    makeToken({
      surface: 'は',
      startPos: 0,
      endPos: 1,
      pos1: '助詞',
      pos2: '係助詞',
      pos3: '一般',
    }),
  ];

  const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
  assert.equal(enriched[0]?.pos2, '係助詞');
  assert.equal(enriched[0]?.pos3, '一般');
});

test('enrichTokensWithMecabPos1 keeps partOfSpeech unchanged and only enriches POS tags', () => {
  const tokens = [makeToken({ surface: 'これは', startPos: 0, endPos: 3 })];
  const mecabTokens = [
    makeToken({
      surface: 'これ',
      startPos: 0,
      endPos: 2,
      pos1: '名詞',
      partOfSpeech: PartOfSpeech.noun,
    }),
    makeToken({
      surface: 'は',
      startPos: 2,
      endPos: 3,
      pos1: '助詞',
      partOfSpeech: PartOfSpeech.particle,
    }),
  ];

  const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
  assert.equal(enriched[0]?.pos1, '名詞|助詞');
  assert.equal(enriched[0]?.partOfSpeech, PartOfSpeech.other);
});

test('enrichTokensWithMecabPos1 passes through unchanged when mecab tokens are null or empty', () => {
  const tokens = [makeToken({ surface: '猫', startPos: 0, endPos: 1 })];

  const nullResult = enrichTokensWithMecabPos1(tokens, null);
  assert.strictEqual(nullResult, tokens);

  const emptyResult = enrichTokensWithMecabPos1(tokens, []);
  assert.strictEqual(emptyResult, tokens);
});

test('enrichTokensWithMecabPos1 avoids repeated full scans over distant mecab surfaces', () => {
  const tokens = Array.from({ length: 12 }, (_, index) =>
    makeToken({ surface: `w${index}`, startPos: index, endPos: index + 1, pos1: '' }),
  );
  const mecabTokens = tokens.map((token) =>
    makeToken({
      surface: token.surface,
      startPos: token.startPos,
      endPos: token.endPos,
      pos1: '名詞',
    }),
  );

  let distantSurfaceReads = 0;
  const distantToken = makeToken({ surface: '遠', startPos: 500, endPos: 501, pos1: '記号' });
  Object.defineProperty(distantToken, 'surface', {
    configurable: true,
    get() {
      distantSurfaceReads += 1;
      if (distantSurfaceReads > 3) {
        throw new Error('repeated full scan detected');
      }
      return '遠';
    },
  });
  mecabTokens.push(distantToken);

  const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
  assert.equal(enriched.length, tokens.length);
  for (const token of enriched) {
    assert.equal(token.pos1, '名詞');
  }
});

test('enrichTokensWithMecabPos1 avoids repeated active-candidate filter scans', () => {
  const tokens = Array.from({ length: 8 }, (_, index) =>
    makeToken({ surface: `u${index}`, startPos: index, endPos: index + 1, pos1: '' }),
  );
  const mecabTokens = [
    makeToken({ surface: 'SENTINEL', startPos: 0, endPos: 100, pos1: '記号' }),
    ...tokens.map((token, index) =>
      makeToken({
        surface: `m${index}`,
        startPos: token.startPos,
        endPos: token.endPos,
        pos1: '名詞',
      }),
    ),
  ];

  let sentinelFilterCalls = 0;
  const originalFilter = Array.prototype.filter;
  Array.prototype.filter = function filterWithSentinelCheck(
    this: unknown[],
    ...args: any[]
  ): any[] {
    const target = this as Array<{ surface?: string }>;
    if (target.some((candidate) => candidate?.surface === 'SENTINEL')) {
      sentinelFilterCalls += 1;
      if (sentinelFilterCalls > 2) {
        throw new Error('repeated active candidate filter scan detected');
      }
    }
    return (originalFilter as (...params: any[]) => any[]).apply(this, args);
  } as typeof Array.prototype.filter;

  try {
    const enriched = enrichTokensWithMecabPos1(tokens, mecabTokens);
    assert.equal(enriched.length, tokens.length);
  } finally {
    Array.prototype.filter = originalFilter;
  }
});
