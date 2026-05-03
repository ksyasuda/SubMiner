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
      [{ text: '俺は公園にいきたい', reading: 'おれはこうえんにいきたい', headword: '行きたい' }],
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

test('returns null when scanning parser candidates have no dictionary headwords', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [{ text: 'これは', reading: 'これは' }],
      [{ text: 'テスト', reading: 'てすと' }],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.equal(tokens, null);
});

test('drops scanning parser tokens which have no dictionary headword', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [{ text: '(ダクネスの荒い息)', reading: 'だくねすのあらいいき' }],
      [{ text: 'アクア', reading: 'あくあ', headword: 'アクア' }],
      [{ text: 'トラウマ', reading: 'とらうま', headword: 'トラウマ' }],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.deepEqual(
    tokens?.map((token) => ({ surface: token.surface, headword: token.headword })),
    [
      { surface: 'アクア', headword: 'アクア' },
      { surface: 'トラウマ', headword: 'トラウマ' },
    ],
  );
});

test('prefers the longest dictionary headword across merged segments', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [
        { text: 'バニ', reading: 'ばに', headword: 'バニ' },
        { text: 'ール', reading: 'ーる', headword: 'バニール' },
      ],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.deepEqual(
    tokens?.map((token) => ({
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
    })),
    [
      {
        surface: 'バニール',
        reading: 'ばにーる',
        headword: 'バニール',
      },
    ],
  );
});

test('splits trailing grammar endings when later segments are standalone words', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [
        { text: '猫', reading: 'ねこ', headword: '猫' },
        { text: 'です', reading: 'です', headword: 'です' },
      ],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.deepEqual(
    tokens?.map((token) => ({
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
    })),
    [
      {
        surface: '猫',
        reading: 'ねこ',
        headword: '猫',
      },
      {
        surface: 'です',
        reading: 'です',
        headword: 'です',
      },
    ],
  );
});

test('splits trailing ja-nai grammar endings from preceding content', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [
        { text: 'いる', reading: 'いる', headword: 'いる' },
        { text: 'じゃない', reading: 'じゃない', headword: 'じゃない' },
      ],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.deepEqual(
    tokens?.map((token) => ({
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
    })),
    [
      {
        surface: 'いる',
        reading: 'いる',
        headword: 'いる',
      },
      {
        surface: 'じゃない',
        reading: 'じゃない',
        headword: 'じゃない',
      },
    ],
  );
});

test('merges trailing katakana continuation without headword into previous token', () => {
  const parseResults = [
    makeParseItem('scanning-parser', [
      [{ text: 'カズ', reading: 'かず', headword: 'カズマ' }],
      [{ text: 'マ', reading: 'ま' }],
      [{ text: '魔王軍', reading: 'まおうぐん', headword: '魔王軍' }],
    ]),
  ];

  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  assert.deepEqual(
    tokens?.map((token) => ({
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
    })),
    [
      {
        surface: 'カズマ',
        reading: 'かずま',
        headword: 'カズマ',
      },
      {
        surface: '魔王軍',
        reading: 'まおうぐん',
        headword: '魔王軍',
      },
    ],
  );
});

// Regression: merged content+function token candidate must not beat a multi-token split
// candidate that preserves the content token as a standalone frequency-eligible unit.
// Background: Yomitan scanning can produce a single-token candidate where a content word
// is merged with trailing function particles (e.g. かかってこいよ → headword かかってくる).
// When a competing multi-token candidate splits content and function separately, the
// multi-token candidate should win so the content token remains frequency-highlightable.
test('multi-token candidate beats single merged content+function token candidate (frequency regression)', () => {
  // Candidate A: single merged token — content verb fused with trailing sentence-final particle
  // This is the "bad" candidate: downstream annotation would exclude frequency for the whole
  // token because the merged pos1 would contain a function-word component.
  const mergedCandidate = makeParseItem('scanning-parser', [
    [{ text: 'かかってこいよ', reading: 'かかってこいよ', headword: 'かかってくる' }],
  ]);

  // Candidate B: two tokens — content verb surface + particle separately.
  // The content token is frequency-eligible on its own.
  const splitCandidate = makeParseItem('scanning-parser', [
    [{ text: 'かかってこい', reading: 'かかってこい', headword: 'かかってくる' }],
    [{ text: 'よ', reading: 'よ', headword: 'よ' }],
  ]);

  // When merged candidate comes first in the array, multi-token split still wins.
  const tokens = selectYomitanParseTokens(
    [mergedCandidate, splitCandidate],
    () => false,
    'headword',
  );
  assert.equal(tokens?.length, 2);
  assert.equal(tokens?.[0]?.surface, 'かかってこい');
  assert.equal(tokens?.[0]?.headword, 'かかってくる');
  assert.equal(tokens?.[1]?.surface, 'よ');
});

test('multi-token candidate beats single merged content+function token regardless of input order', () => {
  const mergedCandidate = makeParseItem('scanning-parser', [
    [{ text: 'かかってこいよ', reading: 'かかってこいよ', headword: 'かかってくる' }],
  ]);

  const splitCandidate = makeParseItem('scanning-parser', [
    [{ text: 'かかってこい', reading: 'かかってこい', headword: 'かかってくる' }],
    [{ text: 'よ', reading: 'よ', headword: 'よ' }],
  ]);

  // Split candidate comes first — should still win over merged.
  const tokens = selectYomitanParseTokens(
    [splitCandidate, mergedCandidate],
    () => false,
    'headword',
  );
  assert.equal(tokens?.length, 2);
  assert.equal(tokens?.[0]?.surface, 'かかってこい');
  assert.equal(tokens?.[1]?.surface, 'よ');
});
