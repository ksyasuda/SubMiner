import test from 'node:test';
import assert from 'node:assert/strict';
import { PartOfSpeech } from '../../types';
import { createTokenizerDepsRuntime, TokenizerServiceDeps, tokenizeSubtitle } from './tokenizer';

function makeDeps(overrides: Partial<TokenizerServiceDeps> = {}): TokenizerServiceDeps {
  return {
    getYomitanExt: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    isKnownWord: () => false,
    getKnownWordMatchMode: () => 'headword',
    getJlptLevel: () => null,
    tokenizeWithMecab: async () => null,
    ...overrides,
  };
}

interface YomitanTokenInput {
  surface: string;
  reading?: string;
  headword?: string;
}

function makeDepsFromYomitanTokens(
  tokens: YomitanTokenInput[],
  overrides: Partial<TokenizerServiceDeps> = {},
): TokenizerServiceDeps {
  return makeDeps({
    getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
    getYomitanParserWindow: () =>
      ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: 'scanning-parser',
              index: 0,
              content: tokens.map((token) => [
                {
                  text: token.surface,
                  reading: token.reading ?? token.surface,
                  headwords: [[{ term: token.headword ?? token.surface }]],
                },
              ]),
            },
          ],
        },
      }) as unknown as Electron.BrowserWindow,
    ...overrides,
  });
}

test('tokenizeSubtitle assigns JLPT level to parsed Yomitan tokens', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '猫',
                      reading: 'ねこ',
                      headwords: [[{ term: '猫' }]],
                    },
                    {
                      text: 'です',
                      reading: 'です',
                      headwords: [[{ term: 'です' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => null,
      getJlptLevel: (text) => (text === '猫' ? 'N5' : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, 'N5');
});

test('tokenizeSubtitle caches JLPT lookups across repeated tokens', async () => {
  let lookupCalls = 0;
  const result = await tokenizeSubtitle(
    '猫猫',
    makeDepsFromYomitanTokens(
      [
        { surface: '猫', reading: 'ねこ', headword: '猫' },
        { surface: '猫', reading: 'ねこ', headword: '猫' },
      ],
      {
        getJlptLevel: (text) => {
          lookupCalls += 1;
          return text === '猫' ? 'N5' : null;
        },
      },
    ),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(lookupCalls, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, 'N5');
  assert.equal(result.tokens?.[1]?.jlptLevel, 'N5');
});

test('tokenizeSubtitle leaves JLPT unset for non-matching tokens', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      getJlptLevel: () => null,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test('tokenizeSubtitle skips JLPT lookups when disabled', async () => {
  let lookupCalls = 0;
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      getJlptLevel: () => {
        lookupCalls += 1;
        return 'N5';
      },
      getJlptEnabled: () => false,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
  assert.equal(lookupCalls, 0);
});

test('tokenizeSubtitle applies frequency dictionary ranks', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens(
      [
        { surface: '猫', reading: 'ねこ', headword: '猫' },
        { surface: 'です', reading: 'です', headword: 'です' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '猫' ? 23 : 1200),
      },
    ),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.frequencyRank, 23);
  assert.equal(result.tokens?.[1]?.frequencyRank, 1200);
});

test('tokenizeSubtitle uses only selected Yomitan headword for frequency lookup', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '猫です',
                      reading: 'ねこです',
                      headwords: [[{ term: '猫です' }], [{ term: '猫' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyRank: (text) => (text === '猫' ? 40 : text === '猫です' ? 1200 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 1200);
});

test('tokenizeSubtitle keeps furigana-split Yomitan segments as one token', async () => {
  const result = await tokenizeSubtitle(
    '友達と話した',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '友',
                      reading: 'とも',
                      headwords: [[{ term: '友達' }]],
                    },
                    {
                      text: '達',
                      reading: 'だち',
                    },
                  ],
                  [
                    {
                      text: 'と',
                      reading: 'と',
                      headwords: [[{ term: 'と' }]],
                    },
                  ],
                  [
                    {
                      text: '話した',
                      reading: 'はなした',
                      headwords: [[{ term: '話す' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyRank: (text) => (text === '友達' ? 22 : text === '話す' ? 90 : null),
    }),
  );

  assert.equal(result.tokens?.length, 3);
  assert.equal(result.tokens?.[0]?.surface, '友達');
  assert.equal(result.tokens?.[0]?.reading, 'ともだち');
  assert.equal(result.tokens?.[0]?.headword, '友達');
  assert.equal(result.tokens?.[0]?.frequencyRank, 22);
  assert.equal(result.tokens?.[1]?.surface, 'と');
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[2]?.surface, '話した');
  assert.equal(result.tokens?.[2]?.frequencyRank, 90);
});

test('tokenizeSubtitle prefers exact headword frequency over surface/reading when available', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '猫',
                      reading: 'ねこ',
                      headwords: [[{ term: 'ネコ' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyRank: (text) => (text === '猫' ? 1200 : text === 'ネコ' ? 8 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 8);
});

test('tokenizeSubtitle keeps no frequency when only reading matches and headword misses', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '猫',
                      reading: 'ねこ',
                      headwords: [[{ term: '猫です' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyRank: (text) => (text === 'ねこ' ? 77 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle ignores invalid frequency rank on selected headword', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '猫です',
                      reading: 'ねこです',
                      headwords: [[{ term: '猫' }], [{ term: '猫です' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyRank: (text) => (text === '猫' ? Number.NaN : text === '猫です' ? 500 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle handles real-word frequency candidates and prefers most frequent term', async () => {
  const result = await tokenizeSubtitle(
    '昨日',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '昨日',
                      reading: 'きのう',
                      headwords: [[{ term: '昨日' }], [{ term: 'きのう' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyRank: (text) => (text === 'きのう' ? 120 : text === '昨日' ? 40 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 40);
});

test('tokenizeSubtitle ignores candidates with no dictionary rank when higher-frequency candidate exists', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '猫',
                      reading: 'ねこ',
                      headwords: [
                        [{ term: '猫' }],
                        [{ term: '猫です' }],
                        [{ term: 'unknown-term' }],
                      ],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyRank: (text) =>
        text === 'unknown-term' ? -1 : text === '猫' ? 88 : text === '猫です' ? 9000 : null,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 88);
});

test('tokenizeSubtitle ignores frequency lookup failures', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      tokenizeWithMecab: async () => [
        {
          headword: '猫',
          surface: '猫',
          reading: 'ネコ',
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getFrequencyRank: () => {
        throw new Error('frequency lookup unavailable');
      },
    }),
  );

  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle skips frequency rank when Yomitan token is enriched as particle by mecab pos1', async () => {
  const result = await tokenizeSubtitle(
    'は',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: 'は',
                      reading: 'は',
                      headwords: [[{ term: 'は' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => [
        {
          headword: 'は',
          surface: 'は',
          reading: 'ハ',
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getFrequencyRank: (text) => (text === 'は' ? 10 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.pos1, '助詞');
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle ignores invalid frequency ranks', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromYomitanTokens(
      [
        { surface: '猫', reading: 'ねこ', headword: '猫' },
        { surface: 'です', reading: 'です', headword: 'です' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => {
          if (text === '猫') return Number.NaN;
          if (text === 'です') return -1;
          return 100;
        },
      },
    ),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
});

test('tokenizeSubtitle skips frequency lookups when disabled', async () => {
  let frequencyCalls = 0;
  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      getFrequencyDictionaryEnabled: () => false,
      getFrequencyRank: () => {
        frequencyCalls += 1;
        return 10;
      },
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(frequencyCalls, 0);
});

test('tokenizeSubtitle skips JLPT level for excluded demonstratives', async () => {
  const result = await tokenizeSubtitle(
    'この',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: 'この',
                      reading: 'この',
                      headwords: [[{ term: 'この' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => null,
      getJlptLevel: (text) => (text === 'この' ? 'N5' : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test('tokenizeSubtitle skips JLPT level for repeated kana SFX', async () => {
  const result = await tokenizeSubtitle(
    'ああ',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: 'ああ',
                      reading: 'ああ',
                      headwords: [[{ term: 'ああ' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => null,
      getJlptLevel: (text) => (text === 'ああ' ? 'N5' : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test('tokenizeSubtitle assigns JLPT level to Yomitan tokens', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      getJlptLevel: (text) => (text === '猫' ? 'N4' : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, 'N4');
});

test('tokenizeSubtitle can assign JLPT level to Yomitan particle token', async () => {
  const result = await tokenizeSubtitle(
    'は',
    makeDepsFromYomitanTokens([{ surface: 'は', reading: 'は', headword: 'は' }], {
      getJlptLevel: (text) => (text === 'は' ? 'N5' : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, 'N5');
});

test('tokenizeSubtitle returns null tokens for empty normalized text', async () => {
  const result = await tokenizeSubtitle(' \\n  ', makeDeps());
  assert.deepEqual(result, { text: ' \\n  ', tokens: null });
});

test('tokenizeSubtitle normalizes newlines before Yomitan parse request', async () => {
  let parseInput = '';
  const result = await tokenizeSubtitle(
    '猫\\Nです\nね',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              parseInput = script;
              return null;
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.match(parseInput, /猫 です ね/);
  assert.equal(result.text, '猫\nです\nね');
  assert.equal(result.tokens, null);
});

test('tokenizeSubtitle returns null tokens when Yomitan parsing is unavailable', async () => {
  const result = await tokenizeSubtitle('猫です', makeDeps());

  assert.deepEqual(result, { text: '猫です', tokens: null });
});

test('tokenizeSubtitle returns null tokens when mecab throws', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      tokenizeWithMecab: async () => {
        throw new Error('mecab failed');
      },
    }),
  );

  assert.deepEqual(result, { text: '猫です', tokens: null });
});

test('tokenizeSubtitle uses Yomitan parser result when available', async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async () => [
        {
          source: 'scanning-parser',
          index: 0,
          content: [
            [
              {
                text: '猫',
                reading: 'ねこ',
                headwords: [[{ term: '猫' }]],
              },
            ],
            [
              {
                text: 'です',
                reading: 'です',
              },
            ],
          ],
        },
      ],
    },
  } as unknown as Electron.BrowserWindow;

  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () => parserWindow,
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, '猫');
  assert.equal(result.tokens?.[0]?.reading, 'ねこ');
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.surface, 'です');
  assert.equal(result.tokens?.[1]?.reading, 'です');
  assert.equal(result.tokens?.[1]?.isKnown, false);
});

test('tokenizeSubtitle logs selected Yomitan groups when debug toggle is enabled', async () => {
  const infoLogs: string[] = [];
  const originalInfo = console.info;
  console.info = (...args: unknown[]) => {
    infoLogs.push(args.map((value) => String(value)).join(' '));
  };

  try {
    await tokenizeSubtitle(
      '友達と話した',
      makeDeps({
        getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
        getYomitanParserWindow: () =>
          ({
            isDestroyed: () => false,
            webContents: {
              executeJavaScript: async () => [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [
                      {
                        text: '友',
                        reading: 'とも',
                        headwords: [[{ term: '友達' }]],
                      },
                      {
                        text: '達',
                        reading: 'だち',
                      },
                    ],
                    [
                      {
                        text: 'と',
                        reading: 'と',
                        headwords: [[{ term: 'と' }]],
                      },
                    ],
                  ],
                },
              ],
            },
          }) as unknown as Electron.BrowserWindow,
        tokenizeWithMecab: async () => null,
        getYomitanGroupDebugEnabled: () => true,
      }),
    );
  } finally {
    console.info = originalInfo;
  }

  assert.ok(infoLogs.some((line) => line.includes('Selected Yomitan token groups')));
});

test('tokenizeSubtitle does not log Yomitan groups when debug toggle is disabled', async () => {
  const infoLogs: string[] = [];
  const originalInfo = console.info;
  console.info = (...args: unknown[]) => {
    infoLogs.push(args.map((value) => String(value)).join(' '));
  };

  try {
    await tokenizeSubtitle(
      '友達と話した',
      makeDeps({
        getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
        getYomitanParserWindow: () =>
          ({
            isDestroyed: () => false,
            webContents: {
              executeJavaScript: async () => [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [
                      {
                        text: '友',
                        reading: 'とも',
                        headwords: [[{ term: '友達' }]],
                      },
                      {
                        text: '達',
                        reading: 'だち',
                      },
                    ],
                  ],
                },
              ],
            },
          }) as unknown as Electron.BrowserWindow,
        tokenizeWithMecab: async () => null,
        getYomitanGroupDebugEnabled: () => false,
      }),
    );
  } finally {
    console.info = originalInfo;
  }

  assert.equal(
    infoLogs.some((line) => line.includes('Selected Yomitan token groups')),
    false,
  );
});

test('tokenizeSubtitle preserves segmented Yomitan line as one token', async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async () => [
        {
          source: 'scanning-parser',
          index: 0,
          content: [
            [
              {
                text: '猫',
                reading: 'ねこ',
                headwords: [[{ term: '猫です' }]],
              },
              {
                text: 'です',
                reading: 'です',
              },
            ],
          ],
        },
      ],
    },
  } as unknown as Electron.BrowserWindow;

  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () => parserWindow,
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '猫です');
  assert.equal(result.tokens?.[0]?.reading, 'ねこです');
  assert.equal(result.tokens?.[0]?.headword, '猫です');
  assert.equal(result.tokens?.[0]?.isKnown, false);
});

test('tokenizeSubtitle keeps scanning parser token when scanning parser returns one token', async () => {
  const result = await tokenizeSubtitle(
    '俺は小園にいきたい',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '俺は小園にいきたい',
                      reading: 'おれは小園にいきたい',
                      headwords: [[{ term: '俺は小園にいきたい' }]],
                    },
                  ],
                ],
              },
              {
                source: 'mecab',
                index: 0,
                content: [
                  [
                    {
                      text: '俺',
                      reading: 'おれ',
                      headwords: [[{ term: '俺' }]],
                    },
                  ],
                  [
                    {
                      text: 'は',
                      reading: 'は',
                      headwords: [[{ term: 'は' }]],
                    },
                  ],
                  [
                    {
                      text: '小園',
                      reading: 'おうえん',
                      headwords: [[{ term: '小園' }]],
                    },
                  ],
                  [
                    {
                      text: 'に',
                      reading: 'に',
                      headwords: [[{ term: 'に' }]],
                    },
                  ],
                  [
                    {
                      text: 'いきたい',
                      reading: 'いきたい',
                      headwords: [[{ term: 'いきたい' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyDictionaryEnabled: () => true,
      tokenizeWithMecab: async () => null,
      getFrequencyRank: (text) => (text === '小園' ? 25 : text === 'いきたい' ? 1500 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.map((token) => token.surface).join(','), '俺は小園にいきたい');
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle keeps scanning parser tokens when they are already split', async () => {
  const result = await tokenizeSubtitle(
    '小園に行きたい',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '小園',
                      reading: 'おうえん',
                      headwords: [[{ term: '小園' }]],
                    },
                  ],
                  [
                    {
                      text: 'に',
                      reading: 'に',
                      headwords: [[{ term: 'に' }]],
                    },
                  ],
                  [
                    {
                      text: '行きたい',
                      reading: 'いきたい',
                      headwords: [[{ term: '行きたい' }]],
                    },
                  ],
                ],
              },
              {
                source: 'mecab',
                index: 0,
                content: [
                  [
                    {
                      text: '小',
                      reading: 'お',
                      headwords: [[{ term: '小' }]],
                    },
                  ],
                  [
                    {
                      text: '園',
                      reading: 'えん',
                      headwords: [[{ term: '園' }]],
                    },
                  ],
                  [
                    {
                      text: 'に',
                      reading: 'に',
                      headwords: [[{ term: 'に' }]],
                    },
                  ],
                  [
                    {
                      text: '行き',
                      reading: 'いき',
                      headwords: [[{ term: '行き' }]],
                    },
                  ],
                  [
                    {
                      text: 'たい',
                      reading: 'たい',
                      headwords: [[{ term: 'たい' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '小園' ? 20 : null),
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.tokens?.length, 3);
  assert.equal(result.tokens?.map((token) => token.surface).join(','), '小園,に,行きたい');
  assert.equal(result.tokens?.[0]?.frequencyRank, 20);
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[2]?.frequencyRank, undefined);
});

test('tokenizeSubtitle keeps parsing explicit by scanning-parser source only', async () => {
  const result = await tokenizeSubtitle(
    '俺は公園にいきたい',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '俺',
                      reading: 'おれ',
                      headwords: [[{ term: '俺' }]],
                    },
                  ],
                  [{ text: 'は', reading: '', headwords: [[{ term: 'は' }]] }],
                  [
                    {
                      text: '公園',
                      reading: 'こうえん',
                      headwords: [[{ term: '公園' }]],
                    },
                  ],
                  [
                    {
                      text: 'にい',
                      reading: '',
                      headwords: [[{ term: '兄' }], [{ term: '二位' }]],
                    },
                  ],
                  [
                    {
                      text: 'きたい',
                      reading: '',
                      headwords: [[{ term: '期待' }], [{ term: '来る' }]],
                    },
                  ],
                ],
              },
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '俺',
                      reading: 'おれ',
                      headwords: [[{ term: '俺' }]],
                    },
                  ],
                  [
                    {
                      text: 'は',
                      reading: 'は',
                      headwords: [[{ term: 'は' }]],
                    },
                  ],
                  [
                    {
                      text: '公園',
                      reading: 'こうえん',
                      headwords: [[{ term: '公園' }]],
                    },
                  ],
                  [
                    {
                      text: 'に',
                      reading: 'に',
                      headwords: [[{ term: 'に' }]],
                    },
                  ],
                  [
                    {
                      text: '行きたい',
                      reading: 'いきたい',
                      headwords: [[{ term: '行きたい' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) =>
        text === '俺' ? 51 : text === '公園' ? 2304 : text === '行きたい' ? 1500 : null,
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.tokens?.map((token) => token.surface).join(','), '俺,は,公園,に,行きたい');
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[3]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[4]?.frequencyRank, 1500);
});

test('tokenizeSubtitle still assigns frequency to non-known Yomitan tokens', async () => {
  const result = await tokenizeSubtitle(
    '小園に',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async () => [
              {
                source: 'scanning-parser',
                index: 0,
                content: [
                  [
                    {
                      text: '小園',
                      reading: 'おうえん',
                      headwords: [[{ term: '小園' }]],
                    },
                  ],
                  [
                    {
                      text: 'に',
                      reading: 'に',
                      headwords: [[{ term: 'に' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '小園' ? 75 : text === 'に' ? 3000 : null),
      isKnownWord: (text) => text === '小園',
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.isKnown, true);
  assert.equal(result.tokens?.[0]?.frequencyRank, 75);
  assert.equal(result.tokens?.[1]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.frequencyRank, 3000);
});

test('tokenizeSubtitle marks tokens as known using callback', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      isKnownWord: (text) => text === '猫',
    }),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.[0]?.isKnown, true);
});

test('tokenizeSubtitle still assigns frequency rank to non-known tokens', async () => {
  const result = await tokenizeSubtitle(
    '既知未知',
    makeDepsFromYomitanTokens(
      [
        { surface: '既知', reading: 'きち', headword: '既知' },
        { surface: '未知', reading: 'みち', headword: '未知' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '既知' ? 20 : text === '未知' ? 30 : null),
        isKnownWord: (text) => text === '既知',
      },
    ),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.isKnown, true);
  assert.equal(result.tokens?.[0]?.frequencyRank, 20);
  assert.equal(result.tokens?.[1]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.frequencyRank, 30);
});

test('tokenizeSubtitle selects one N+1 target token', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens(
      [
        { surface: '私', reading: 'わたし', headword: '私' },
        { surface: '犬', reading: 'いぬ', headword: '犬' },
      ],
      {
        getMinSentenceWordsForNPlusOne: () => 2,
        isKnownWord: (text) => text === '私',
      },
    ),
  );

  const targets = result.tokens?.filter((token) => token.isNPlusOneTarget) ?? [];
  assert.equal(targets.length, 1);
  assert.equal(targets[0]?.surface, '犬');
});

test('tokenizeSubtitle does not mark target when sentence has multiple candidates', async () => {
  const result = await tokenizeSubtitle(
    '猫犬',
    makeDepsFromYomitanTokens(
      [
        { surface: '猫', reading: 'ねこ', headword: '猫' },
        { surface: '犬', reading: 'いぬ', headword: '犬' },
      ],
      {},
    ),
  );

  assert.equal(
    result.tokens?.some((token) => token.isNPlusOneTarget),
    false,
  );
});

test('tokenizeSubtitle applies N+1 target marking to Yomitan results', async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async () => [
        {
          source: 'scanning-parser',
          index: 0,
          content: [
            [
              {
                text: '猫',
                reading: 'ねこ',
                headwords: [[{ term: '猫' }]],
              },
            ],
            [
              {
                text: 'です',
                reading: 'です',
                headwords: [[{ term: 'です' }]],
              },
            ],
          ],
        },
      ],
    },
  } as unknown as Electron.BrowserWindow;

  const result = await tokenizeSubtitle(
    '猫です',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () => parserWindow,
      tokenizeWithMecab: async () => null,
      isKnownWord: (text) => text === 'です',
      getMinSentenceWordsForNPlusOne: () => 2,
    }),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, '猫');
  assert.equal(result.tokens?.[0]?.isNPlusOneTarget, true);
  assert.equal(result.tokens?.[1]?.isNPlusOneTarget, false);
});

test('tokenizeSubtitle ignores Yomitan functional tokens when evaluating N+1 candidates', async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async () => [
        {
          source: 'scanning-parser',
          index: 0,
          content: [
            [{ text: '私', reading: 'わたし', headwords: [[{ term: '私' }]] }],
            [{ text: 'も', reading: 'も', headwords: [[{ term: 'も' }]] }],
            [{ text: 'あの', reading: 'あの', headwords: [[{ term: 'あの' }]] }],
            [{ text: '仮面', reading: 'かめん', headwords: [[{ term: '仮面' }]] }],
            [{ text: 'が', reading: 'が', headwords: [[{ term: 'が' }]] }],
            [{ text: '欲しい', reading: 'ほしい', headwords: [[{ term: '欲しい' }]] }],
            [{ text: 'です', reading: 'です', headwords: [[{ term: 'です' }]] }],
          ],
        },
      ],
    },
  } as unknown as Electron.BrowserWindow;

  const result = await tokenizeSubtitle(
    '私も あの仮面が欲しいです',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () => parserWindow,
      tokenizeWithMecab: async () => [
        {
          surface: '私',
          reading: 'ワタシ',
          headword: '私',
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: 'も',
          reading: 'モ',
          headword: 'も',
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: 'あの',
          reading: 'アノ',
          headword: 'あの',
          startPos: 2,
          endPos: 4,
          partOfSpeech: PartOfSpeech.other,
          pos1: '連体詞',
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: '仮面',
          reading: 'カメン',
          headword: '仮面',
          startPos: 4,
          endPos: 6,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: 'が',
          reading: 'ガ',
          headword: 'が',
          startPos: 6,
          endPos: 7,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: '欲しい',
          reading: 'ホシイ',
          headword: '欲しい',
          startPos: 7,
          endPos: 10,
          partOfSpeech: PartOfSpeech.i_adjective,
          pos1: '形容詞',
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: 'です',
          reading: 'デス',
          headword: 'です',
          startPos: 10,
          endPos: 12,
          partOfSpeech: PartOfSpeech.bound_auxiliary,
          pos1: '助動詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      isKnownWord: (text) => text === '私' || text === 'あの' || text === '欲しい',
    }),
  );

  const targets = result.tokens?.filter((token) => token.isNPlusOneTarget) ?? [];
  assert.equal(targets.length, 1);
  assert.equal(targets[0]?.surface, '仮面');
});

test('tokenizeSubtitle keeps correct MeCab pos1 enrichment when Yomitan offsets skip spaces', async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async () => [
        {
          source: 'scanning-parser',
          index: 0,
          content: [
            [{ text: '私', reading: 'わたし', headwords: [[{ term: '私' }]] }],
            [{ text: 'も', reading: 'も', headwords: [[{ term: 'も' }]] }],
            [{ text: 'あの', reading: 'あの', headwords: [[{ term: 'あの' }]] }],
            [{ text: '仮面', reading: 'かめん', headwords: [[{ term: '仮面' }]] }],
            [{ text: 'が', reading: 'が', headwords: [[{ term: 'が' }]] }],
            [{ text: '欲しい', reading: 'ほしい', headwords: [[{ term: '欲しい' }]] }],
            [{ text: 'です', reading: 'です', headwords: [[{ term: 'です' }]] }],
          ],
        },
      ],
    },
  } as unknown as Electron.BrowserWindow;

  const result = await tokenizeSubtitle(
    '私も あの仮面が欲しいです',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () => parserWindow,
      tokenizeWithMecab: async () => [
        {
          surface: '私',
          reading: 'ワタシ',
          headword: '私',
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: 'も',
          reading: 'モ',
          headword: 'も',
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: ' ',
          reading: '',
          headword: ' ',
          startPos: 2,
          endPos: 3,
          partOfSpeech: PartOfSpeech.symbol,
          pos1: '記号',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: 'あの',
          reading: 'アノ',
          headword: 'あの',
          startPos: 3,
          endPos: 5,
          partOfSpeech: PartOfSpeech.other,
          pos1: '連体詞',
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: '仮面',
          reading: 'カメン',
          headword: '仮面',
          startPos: 5,
          endPos: 7,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: 'が',
          reading: 'ガ',
          headword: 'が',
          startPos: 7,
          endPos: 8,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: '欲しい',
          reading: 'ホシイ',
          headword: '欲しい',
          startPos: 8,
          endPos: 11,
          partOfSpeech: PartOfSpeech.i_adjective,
          pos1: '形容詞',
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: 'です',
          reading: 'デス',
          headword: 'です',
          startPos: 11,
          endPos: 13,
          partOfSpeech: PartOfSpeech.bound_auxiliary,
          pos1: '助動詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      isKnownWord: (text) => text === '私' || text === 'あの' || text === '欲しい',
    }),
  );

  const targets = result.tokens?.filter((token) => token.isNPlusOneTarget) ?? [];
  const gaToken = result.tokens?.find((token) => token.surface === 'が');
  const desuToken = result.tokens?.find((token) => token.surface === 'です');
  assert.equal(gaToken?.pos1, '助詞');
  assert.equal(desuToken?.pos1, '助動詞');
  assert.equal(targets.length, 1);
  assert.equal(targets[0]?.surface, '仮面');
});

test('tokenizeSubtitle does not color 1-2 word sentences by default', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens(
      [
        { surface: '私', reading: 'わたし', headword: '私' },
        { surface: '犬', reading: 'いぬ', headword: '犬' },
      ],
      {},
    ),
  );

  assert.equal(
    result.tokens?.some((token) => token.isNPlusOneTarget),
    false,
  );
});

test('tokenizeSubtitle checks known words by headword, not surface', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫です' }], {
      isKnownWord: (text) => text === '猫です',
    }),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.[0]?.isKnown, true);
});

test('tokenizeSubtitle checks known words by surface when configured', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫です' }], {
      getKnownWordMatchMode: () => 'surface',
      isKnownWord: (text) => text === '猫',
    }),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.[0]?.isKnown, true);
});

test('createTokenizerDepsRuntime checks MeCab availability before first tokenizeWithMecab call', async () => {
  let available = false;
  let checkCalls = 0;

  const deps = createTokenizerDepsRuntime({
    getYomitanExt: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    isKnownWord: () => false,
    getKnownWordMatchMode: () => 'headword',
    getJlptLevel: () => null,
    getMecabTokenizer: () => ({
      getStatus: () => ({ available }),
      checkAvailability: async () => {
        checkCalls += 1;
        available = true;
        return true;
      },
      tokenize: async () => {
        if (!available) {
          return null;
        }
        return [
          {
            word: '仮面',
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '一般',
            pos3: '',
            pos4: '',
            inflectionType: '',
            inflectionForm: '',
            headword: '仮面',
            katakanaReading: 'カメン',
            pronunciation: 'カメン',
          },
        ];
      },
    }),
  });

  const first = await deps.tokenizeWithMecab('仮面');
  const second = await deps.tokenizeWithMecab('仮面');

  assert.equal(checkCalls, 1);
  assert.equal(first?.[0]?.surface, '仮面');
  assert.equal(second?.[0]?.surface, '仮面');
});

test('tokenizeSubtitle uses async MeCab enrichment override when provided', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      tokenizeWithMecab: async () => [
        {
          headword: '猫',
          surface: '猫',
          reading: 'ネコ',
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      enrichTokensWithMecab: async (tokens) =>
        tokens.map((token) => ({
          ...token,
          pos1: 'override-pos',
        })),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.pos1, 'override-pos');
});

test('createTokenizerDepsRuntime exposes async MeCab enrichment helper', async () => {
  const deps = createTokenizerDepsRuntime({
    getYomitanExt: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    isKnownWord: () => false,
    getKnownWordMatchMode: () => 'headword',
    getJlptLevel: () => null,
    getMecabTokenizer: () => null,
  });

  const enriched = await deps.enrichTokensWithMecab?.(
    [
      {
        headword: 'は',
        surface: 'は',
        reading: 'は',
        startPos: 0,
        endPos: 1,
        partOfSpeech: PartOfSpeech.other,
        isMerged: true,
        isKnown: false,
        isNPlusOneTarget: false,
      },
    ],
    [
      {
        headword: 'は',
        surface: 'は',
        reading: 'ハ',
        startPos: 0,
        endPos: 1,
        partOfSpeech: PartOfSpeech.particle,
        pos1: '助詞',
        isMerged: false,
        isKnown: false,
        isNPlusOneTarget: false,
      },
    ],
  );

  assert.equal(enriched?.[0]?.pos1, '助詞');
});

test('tokenizeSubtitle skips all enrichment stages when disabled', async () => {
  let knownCalls = 0;
  let mecabCalls = 0;
  let jlptCalls = 0;
  let frequencyCalls = 0;

  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      isKnownWord: () => {
        knownCalls += 1;
        return true;
      },
      getNPlusOneEnabled: () => false,
      getJlptEnabled: () => false,
      getFrequencyDictionaryEnabled: () => false,
      getJlptLevel: () => {
        jlptCalls += 1;
        return 'N5';
      },
      getFrequencyRank: () => {
        frequencyCalls += 1;
        return 10;
      },
      tokenizeWithMecab: async () => {
        mecabCalls += 1;
        return null;
      },
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(result.tokens?.[0]?.isNPlusOneTarget, false);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(knownCalls, 0);
  assert.equal(mecabCalls, 0);
  assert.equal(jlptCalls, 0);
  assert.equal(frequencyCalls, 0);
});

test('tokenizeSubtitle keeps frequency enrichment while n+1 is disabled', async () => {
  let knownCalls = 0;
  let mecabCalls = 0;
  let frequencyCalls = 0;

  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ', headword: '猫' }], {
      isKnownWord: () => {
        knownCalls += 1;
        return true;
      },
      getNPlusOneEnabled: () => false,
      getJlptEnabled: () => false,
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: () => {
        frequencyCalls += 1;
        return 7;
      },
      tokenizeWithMecab: async () => {
        mecabCalls += 1;
        return [
          {
            headword: '猫',
            surface: '猫',
            reading: 'ネコ',
            startPos: 0,
            endPos: 1,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ];
      },
    }),
  );

  assert.equal(result.tokens?.[0]?.frequencyRank, 7);
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(knownCalls, 0);
  assert.equal(mecabCalls, 1);
  assert.equal(frequencyCalls, 1);
});
