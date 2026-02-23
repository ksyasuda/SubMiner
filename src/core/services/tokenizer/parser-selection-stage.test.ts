import test from 'node:test';
import assert from 'node:assert/strict';
import { selectYomitanParseTokens } from './parser-selection-stage';

interface ParseSegmentInput {
  text: string;
  reading?: string;
  headword?: string;
}

function makeParseItem(
  source: string,
  lines: ParseSegmentInput[][],
): {
  source: string;
  index: number;
  content: Array<
    Array<{ text: string; reading?: string; headwords?: Array<Array<{ term: string }>> }>
  >;
} {
  return {
    source,
    index: 0,
    content: lines.map((line) =>
      line.map((segment) => ({
        text: segment.text,
        reading: segment.reading,
        headwords: segment.headword ? [[{ term: segment.headword }]] : undefined,
      })),
    ),
  };
}

test('prefers scanning parser when scanning candidate has more than one token', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [{ text: '小園', reading: 'おうえん', headword: '小園' }],
      [{ text: 'に', reading: 'に', headword: 'に' }],
    ]),
    makeParseItem('mecab', [
      [{ text: '小', reading: 'お', headword: '小' }],
      [{ text: '園', reading: 'えん', headword: '園' }],
      [{ text: 'に', reading: 'に', headword: 'に' }],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.equal(tokens?.map((token) => token.surface).join(','), '小園,に');
});

test('keeps scanning parser candidate when scanning candidate is single token', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [{ text: '俺は公園にいきたい', reading: 'おれはこうえんにいきたい' }],
    ]),
    makeParseItem('mecab', [
      [{ text: '俺', reading: 'おれ', headword: '俺' }],
      [{ text: 'は', reading: 'は', headword: 'は' }],
      [{ text: '公園', reading: 'こうえん', headword: '公園' }],
      [{ text: 'に', reading: 'に', headword: 'に' }],
      [{ text: 'いきたい', reading: 'いきたい', headword: '行きたい' }],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.equal(tokens?.map((token) => token.surface).join(','), '俺は公園にいきたい');
});

test('tie-break prefers fewer suspicious kana fragments', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [{ text: '俺', reading: 'おれ', headword: '俺' }],
      [{ text: 'にい', reading: '', headword: '兄' }],
      [{ text: 'きたい', reading: '', headword: '期待' }],
    ]),
    makeParseItem('scanning-parser', [
      [{ text: '俺', reading: 'おれ', headword: '俺' }],
      [{ text: 'に', reading: 'に', headword: 'に' }],
      [{ text: '行きたい', reading: 'いきたい', headword: '行きたい' }],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.equal(tokens?.map((token) => token.surface).join(','), '俺,に,行きたい');
});

test('returns null when only mecab-source candidates are present', () => {
  const parseResults = [
    makeParseItem('mecab', [
      [{ text: '俺', reading: 'おれ', headword: '俺' }],
      [{ text: 'は', reading: 'は', headword: 'は' }],
      [{ text: '公園', reading: 'こうえん', headword: '公園' }],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.equal(tokens, null);
});
