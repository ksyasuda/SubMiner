import assert from 'node:assert/strict';
import test from 'node:test';
import { KnownWordMaturityTier, MergedToken, PartOfSpeech } from '../../../types';
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

test('annotateTokens attaches maturity tier to known tokens', () => {
  const tokens = [makeToken({ surface: '食べた', headword: '食べる', reading: 'タベタ' })];
  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '食べる',
      getKnownWordTier: (text) => (text === '食べる' ? 'mature' : null),
    }),
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.knownMaturity, 'mature');
});

test('annotateTokens leaves maturity undefined without a tier lookup dep', () => {
  const tokens = [makeToken()];
  const result = annotateTokens(tokens, makeDeps({ isKnownWord: () => true }));

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.knownMaturity, undefined);
});

test('annotateTokens leaves maturity undefined when the tier lookup has no data', () => {
  const tokens = [makeToken()];
  const result = annotateTokens(
    tokens,
    makeDeps({ isKnownWord: () => true, getKnownWordTier: () => null }),
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.knownMaturity, undefined);
});

test('annotateTokens never attaches maturity to unknown tokens', () => {
  const tokens = [makeToken()];
  const result = annotateTokens(
    tokens,
    makeDeps({ isKnownWord: () => false, getKnownWordTier: () => 'mature' }),
  );

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[0]?.knownMaturity, undefined);
});

test('annotateTokens resolves maturity through the kana reading fallback', () => {
  // Token 大体 with a card mined in kana (だいたい): known status comes from
  // the reading fallback, so the tier must follow the same path.
  const tierByText = new Map<string, KnownWordMaturityTier>([['だいたい', 'young']]);
  const seenTierLookups: Array<{
    text: string;
    reading: string | undefined;
    allowReadingOnlyMatch: boolean | undefined;
  }> = [];
  const tokens = [
    makeToken({ surface: '大体', headword: '大体', reading: 'だいたい', endPos: 2 }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'だいたい',
      getKnownWordTier: (text, reading, options) => {
        seenTierLookups.push({
          text,
          reading,
          allowReadingOnlyMatch: options?.allowReadingOnlyMatch,
        });
        return tierByText.get(text) ?? null;
      },
    }),
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.knownMaturity, 'young');
  // The fallback lookup must opt out of reading-only matching, exactly like
  // the boolean known-status fallback.
  const fallbackLookup = seenTierLookups.find((lookup) => lookup.text === 'だいたい');
  assert.equal(fallbackLookup?.allowReadingOnlyMatch, false);
});

test('annotateTokens keeps maturity on POS-excluded known tokens', () => {
  // Known-word annotations survive the POS noise filter; the tier must too.
  const tokens = [makeToken({ surface: 'の', headword: 'の', reading: 'ノ', pos1: '助詞' })];
  const result = annotateTokens(
    tokens,
    makeDeps({ isKnownWord: () => true, getKnownWordTier: () => 'learning' }),
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.knownMaturity, 'learning');
});

test('annotateTokens strips maturity when known-word annotation is disabled', () => {
  const tokens = [makeToken({ knownMaturity: 'mature' })];
  const result = annotateTokens(
    tokens,
    makeDeps({ isKnownWord: () => true, getKnownWordTier: () => 'mature' }),
    { knownWordsEnabled: false },
  );

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[0]?.knownMaturity, undefined);
});
