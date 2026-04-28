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
  isNameMatch?: boolean;
  wordClasses?: string[];
}

function makeDepsFromYomitanTokens(
  tokens: YomitanTokenInput[],
  overrides: Partial<TokenizerServiceDeps> = {},
): TokenizerServiceDeps {
  let cursor = 0;
  return makeDeps({
    getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
    getYomitanParserWindow: () =>
      ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async (script: string) => {
            if (script.includes('getTermFrequencies')) {
              return [];
            }

            cursor = 0;
            return tokens.map((token) => {
              const startPos = cursor;
              const endPos = startPos + token.surface.length;
              cursor = endPos;
              return {
                surface: token.surface,
                reading: token.reading ?? token.surface,
                headword: token.headword ?? token.surface,
                startPos,
                endPos,
                isNameMatch: token.isNameMatch ?? false,
                wordClasses: token.wordClasses,
              };
            });
          },
        },
      }) as unknown as Electron.BrowserWindow,
    ...overrides,
  });
}

function createDeferred<T>() {
  let resolve: ((value: T) => void) | null = null;
  const promise = new Promise<T>((innerResolve) => {
    resolve = innerResolve;
  });
  return {
    promise,
    resolve: (value: T) => {
      resolve?.(value);
    },
  };
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

test('tokenizeSubtitle preserves Yomitan name-match metadata on tokens', async () => {
  const result = await tokenizeSubtitle(
    'アクアです',
    makeDepsFromYomitanTokens([
      { surface: 'アクア', reading: 'あくあ', headword: 'アクア', isNameMatch: true },
      { surface: 'です', reading: 'です', headword: 'です' },
    ]),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal((result.tokens?.[0] as { isNameMatch?: boolean } | undefined)?.isNameMatch, true);
  assert.equal((result.tokens?.[1] as { isNameMatch?: boolean } | undefined)?.isNameMatch, false);
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

test('tokenizeSubtitle uses left-to-right yomitan scanning to keep full katakana name tokens', async () => {
  const result = await tokenizeSubtitle(
    'カズマ 魔王軍',
    makeDeps({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [];
              }

              return [
                {
                  surface: 'カズマ',
                  reading: 'かずま',
                  headword: 'カズマ',
                  startPos: 0,
                  endPos: 3,
                },
                {
                  surface: '魔王軍',
                  reading: 'まおうぐん',
                  headword: '魔王軍',
                  startPos: 4,
                  endPos: 7,
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
      startPos: token.startPos,
      endPos: token.endPos,
    })),
    [
      {
        surface: 'カズマ',
        reading: 'かずま',
        headword: 'カズマ',
        startPos: 0,
        endPos: 3,
      },
      {
        surface: '魔王軍',
        reading: 'まおうぐん',
        headword: '魔王軍',
        startPos: 4,
        endPos: 7,
      },
    ],
  );
});

test('tokenizeSubtitle loads frequency ranks from Yomitan installed dictionaries', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [
                  {
                    term: '猫',
                    reading: 'ねこ',
                    dictionary: 'freq-dict',
                    frequency: 77,
                    displayValue: '77',
                    displayValueParsed: true,
                  },
                ];
              }

              return [
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
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 77);
});

test('tokenizeSubtitle starts Yomitan frequency lookup and MeCab enrichment in parallel', async () => {
  const frequencyDeferred = createDeferred<unknown[]>();
  const mecabDeferred = createDeferred<null>();
  let frequencyRequested = false;
  let mecabRequested = false;

  const pendingResult = tokenizeSubtitle(
    '猫',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                frequencyRequested = true;
                return await frequencyDeferred.promise;
              }

              return [
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
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => {
        mecabRequested = true;
        return await mecabDeferred.promise;
      },
    }),
  );

  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(frequencyRequested, true);
  assert.equal(mecabRequested, true);

  frequencyDeferred.resolve([
    {
      term: '猫',
      reading: 'ねこ',
      dictionary: 'freq-dict',
      frequency: 77,
      displayValue: '77',
      displayValueParsed: true,
    },
  ]);
  mecabDeferred.resolve(null);

  const result = await pendingResult;
  assert.equal(result.tokens?.[0]?.frequencyRank, 77);
});

test('tokenizeSubtitle can signal tokenization-ready before enrichment completes', async () => {
  const frequencyDeferred = createDeferred<unknown[]>();
  const mecabDeferred = createDeferred<null>();
  let tokenizationReadyText: string | null = null;

  const pendingResult = tokenizeSubtitle(
    '猫',
    makeDeps({
      onTokenizationReady: (text) => {
        tokenizationReadyText = text;
      },
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return await frequencyDeferred.promise;
              }

              return [
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
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => {
        return await mecabDeferred.promise;
      },
    }),
  );

  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(tokenizationReadyText, '猫');

  frequencyDeferred.resolve([]);
  mecabDeferred.resolve(null);
  await pendingResult;
});

test('tokenizeSubtitle appends trailing kana to merged Yomitan readings when headword equals surface', async () => {
  const result = await tokenizeSubtitle(
    '断じて見ていない',
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
                    { text: '断', reading: 'だん', headwords: [[{ term: '断じて' }]] },
                    { text: 'じて', reading: '', headwords: [[{ term: 'じて' }]] },
                  ],
                  [
                    { text: '見', reading: 'み', headwords: [[{ term: '見る' }]] },
                    { text: 'ていない', reading: '', headwords: [[{ term: 'ていない' }]] },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, '断じて');
  assert.equal(result.tokens?.[0]?.reading, 'だんじて');
  assert.equal(result.tokens?.[1]?.surface, '見ていない');
  assert.equal(result.tokens?.[1]?.reading, 'み');
});

test('tokenizeSubtitle queries headword frequencies with token reading for disambiguation', async () => {
  const result = await tokenizeSubtitle(
    '鍛えた',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                assert.equal(
                  script.includes('"term":"鍛える","reading":null'),
                  false,
                  'should not eagerly include term-only fallback pair when reading lookup is present',
                );
                if (!script.includes('"term":"鍛える","reading":"きた"')) {
                  return [];
                }
                return [
                  {
                    term: '鍛える',
                    reading: 'きたえる',
                    dictionary: 'freq-dict',
                    frequency: 46961,
                    displayValue: '2847,46961',
                    displayValueParsed: true,
                  },
                ];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [
                      {
                        text: '鍛えた',
                        reading: 'きた',
                        headwords: [[{ term: '鍛える' }]],
                      },
                    ],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.headword, '鍛える');
  assert.equal(result.tokens?.[0]?.reading, 'きた');
  assert.equal(result.tokens?.[0]?.frequencyRank, 2847);
});

test('tokenizeSubtitle falls back to term-only Yomitan frequency lookup when reading is noisy', async () => {
  const result = await tokenizeSubtitle(
    '断じて',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                if (!script.includes('"term":"断じて","reading":null')) {
                  return [];
                }
                return [
                  {
                    term: '断じて',
                    reading: null,
                    dictionary: 'freq-dict',
                    frequency: 7082,
                    displayValue: '7082',
                    displayValueParsed: true,
                  },
                ];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [
                      {
                        text: '断じて',
                        reading: 'だん',
                        headwords: [[{ term: '断じて' }]],
                      },
                    ],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 7082);
});

test('tokenizeSubtitle avoids headword term-only fallback rank when reading-specific frequency exists', async () => {
  const result = await tokenizeSubtitle(
    '無人',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                if (!script.includes('"term":"無人","reading":"むじん"')) {
                  return [];
                }
                return [
                  {
                    term: '無人',
                    reading: null,
                    dictionary: 'CC100',
                    dictionaryPriority: 0,
                    frequency: 157632,
                    displayValue: null,
                    displayValueParsed: false,
                  },
                  {
                    term: '無人',
                    reading: 'むじん',
                    dictionary: 'CC100',
                    dictionaryPriority: 0,
                    frequency: 7141,
                    displayValue: null,
                    displayValueParsed: false,
                  },
                ];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [
                      {
                        text: '無人',
                        reading: 'むじん',
                        headwords: [[{ term: '無人' }]],
                      },
                    ],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 7141);
});

test('tokenizeSubtitle prefers Yomitan frequency from highest-priority dictionary', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [
                  {
                    term: '猫',
                    reading: 'ねこ',
                    dictionary: 'low-priority',
                    dictionaryPriority: 2,
                    frequency: 5,
                    displayValue: '5',
                    displayValueParsed: true,
                  },
                  {
                    term: '猫',
                    reading: 'ねこ',
                    dictionary: 'high-priority',
                    dictionaryPriority: 0,
                    frequency: 100,
                    displayValue: '100',
                    displayValueParsed: true,
                  },
                ];
              }

              return [
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
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 100);
});

test('tokenizeSubtitle ignores occurrence-based Yomitan frequencies for inflected terms', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [
                  {
                    term: '潜む',
                    reading: 'ひそ',
                    dictionary: 'CC100',
                    frequency: 118121,
                    displayValue: null,
                    displayValueParsed: false,
                  },
                ];
              }

              if (script.includes('optionsGetFull')) {
                return {
                  profileCurrent: 0,
                  profileIndex: 0,
                  scanLength: 40,
                  dictionaries: ['CC100'],
                  dictionaryPriorityByName: { CC100: 0 },
                  dictionaryFrequencyModeByName: { CC100: 'occurrence-based' },
                  profiles: [
                    {
                      options: {
                        scanning: { length: 40 },
                        dictionaries: [{ name: 'CC100', enabled: true, id: 0 }],
                      },
                    },
                  ],
                };
              }

              return [
                {
                  surface: '潜み',
                  reading: 'ひそ',
                  headword: '潜む',
                  startPos: 0,
                  endPos: 2,
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle falls back to raw term-only Yomitan rank when no scan-derived rank exists', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [
                  {
                    term: '潜む',
                    reading: 'ひそ',
                    hasReading: false,
                    dictionary: 'CC100',
                    frequency: 118121,
                    displayValue: null,
                    displayValueParsed: false,
                  },
                ];
              }

              if (script.includes('optionsGetFull')) {
                return {
                  profileCurrent: 0,
                  profileIndex: 0,
                  scanLength: 40,
                  dictionaries: ['CC100'],
                  dictionaryPriorityByName: { CC100: 0 },
                  dictionaryFrequencyModeByName: { CC100: 'rank-based' },
                  profiles: [
                    {
                      options: {
                        scanning: { length: 40 },
                        dictionaries: [{ name: 'CC100', enabled: true, id: 0 }],
                      },
                    },
                  ],
                };
              }

              return [
                {
                  surface: '潜み',
                  reading: 'ひそ',
                  headword: '潜む',
                  startPos: 0,
                  endPos: 2,
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 118121);
});

test('tokenizeSubtitle keeps parsed display rank for term-only inflected headword fallback', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [
                  {
                    term: '潜む',
                    reading: 'ひそ',
                    hasReading: false,
                    dictionary: 'CC100',
                    frequency: 118121,
                    displayValue: '118,121',
                    displayValueParsed: false,
                  },
                ];
              }

              if (script.includes('optionsGetFull')) {
                return {
                  profileCurrent: 0,
                  profileIndex: 0,
                  scanLength: 40,
                  dictionaries: ['CC100'],
                  dictionaryPriorityByName: { CC100: 0 },
                  dictionaryFrequencyModeByName: { CC100: 'rank-based' },
                  profiles: [
                    {
                      options: {
                        scanning: { length: 40 },
                        dictionaries: [{ name: 'CC100', enabled: true, id: 0 }],
                      },
                    },
                  ],
                };
              }

              return [
                {
                  surface: '潜み',
                  reading: 'ひそ',
                  headword: '潜む',
                  startPos: 0,
                  endPos: 2,
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 118);
});

test('tokenizeSubtitle preserves scan-derived rank over lower-priority Yomitan fallback', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [
                  {
                    term: '潜む',
                    reading: 'ひそ',
                    hasReading: false,
                    dictionary: 'CC100',
                    dictionaryPriority: 2,
                    frequency: 118121,
                    displayValue: null,
                    displayValueParsed: false,
                  },
                ];
              }

              return [
                {
                  surface: '潜み',
                  reading: 'ひそむ',
                  headword: '潜む',
                  startPos: 0,
                  endPos: 2,
                  frequencyRank: 4073,
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 4073);
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

test('tokenizeSubtitle falls back to exact surface frequency when merged headword lookup misses', async () => {
  const frequencyScripts: string[] = [];
  const result = await tokenizeSubtitle(
    '陰に',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                frequencyScripts.push(script);
                return script.includes('"term":"陰に","reading":"いんに"')
                  ? [
                      {
                        term: '陰に',
                        reading: 'いんに',
                        dictionary: 'freq-dict',
                        frequency: 5702,
                        displayValue: '5702',
                        displayValueParsed: true,
                      },
                    ]
                  : [];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [
                      {
                        text: '陰に',
                        reading: 'いんに',
                        headwords: [[{ term: '陰' }]],
                      },
                    ],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '陰に');
  assert.equal(result.tokens?.[0]?.headword, '陰');
  assert.equal(result.tokens?.[0]?.frequencyRank, 5702);
  assert.equal(
    frequencyScripts.some((script) => script.includes('"term":"陰","reading":"いんに"')),
    true,
  );
  assert.equal(
    frequencyScripts.some((script) => script.includes('"term":"陰に","reading":"いんに"')),
    true,
  );
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

test('tokenizeSubtitle keeps standalone particle token hoverable while clearing annotation metadata', async () => {
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

  assert.equal(result.text, 'は');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
      pos1: token.pos1,
      isKnown: token.isKnown,
      isNPlusOneTarget: token.isNPlusOneTarget,
      isNameMatch: token.isNameMatch,
      jlptLevel: token.jlptLevel,
      frequencyRank: token.frequencyRank,
    })),
    [
      {
        surface: 'は',
        reading: 'は',
        headword: 'は',
        pos1: '助詞',
        isKnown: false,
        isNPlusOneTarget: false,
        isNameMatch: false,
        jlptLevel: undefined,
        frequencyRank: undefined,
      },
    ],
  );
});

test('tokenizeSubtitle keeps frequency rank when mecab tags classify token as content-bearing', async () => {
  const result = await tokenizeSubtitle(
    'ふふ',
    makeDepsFromYomitanTokens([{ surface: 'ふふ', reading: '', headword: 'ふふ' }], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === 'ふふ' ? 3014 : null),
      tokenizeWithMecab: async () => [
        {
          headword: 'ふふ',
          surface: 'ふふ',
          reading: 'フフ',
          startPos: 0,
          endPos: 2,
          partOfSpeech: PartOfSpeech.verb,
          pos1: '動詞',
          pos2: '自立',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 3014);
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

test('tokenizeSubtitle keeps repeated kana interjections tokenized while clearing annotation metadata', async () => {
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

  assert.equal(result.text, 'ああ');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      reading: token.reading,
      jlptLevel: token.jlptLevel,
      frequencyRank: token.frequencyRank,
      isKnown: token.isKnown,
      isNPlusOneTarget: token.isNPlusOneTarget,
    })),
    [
      {
        surface: 'ああ',
        headword: 'ああ',
        reading: 'ああ',
        jlptLevel: undefined,
        frequencyRank: undefined,
        isKnown: false,
        isNPlusOneTarget: false,
      },
    ],
  );
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

test('tokenizeSubtitle clears JLPT level from standalone Yomitan particle token', async () => {
  const result = await tokenizeSubtitle(
    'は',
    makeDepsFromYomitanTokens([{ surface: 'は', reading: 'は', headword: 'は' }], {
      getJlptLevel: (text) => (text === 'は' ? 'N5' : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
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

test('tokenizeSubtitle collapses zero-width separators before Yomitan parse request', async () => {
  let parseInput = '';
  const result = await tokenizeSubtitle(
    'キリキリと\u200bかかってこい\nこのヘナチョコ冒険者どもめが！',
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

  assert.match(parseInput, /キリキリと かかってこい このヘナチョコ冒険者どもめが！/);
  assert.equal(result.text, 'キリキリと\u200bかかってこい\nこのヘナチョコ冒険者どもめが！');
  assert.equal(result.tokens, null);
});

test('tokenizeSubtitle returns null tokens when Yomitan parsing is unavailable', async () => {
  const result = await tokenizeSubtitle('猫です', makeDeps());

  assert.deepEqual(result, { text: '猫です', tokens: null });
});

test('tokenizeSubtitle skips token payload and annotations when Yomitan parse has no dictionary matches', async () => {
  let frequencyRequested = false;
  let jlptLookupCalls = 0;
  let mecabCalls = 0;

  const result = await tokenizeSubtitle(
    'これはテスト',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                frequencyRequested = true;
                return [];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [{ text: 'これは', reading: 'これは' }],
                    [{ text: 'テスト', reading: 'てすと' }],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => {
        mecabCalls += 1;
        return null;
      },
      getJlptLevel: () => {
        jlptLookupCalls += 1;
        return 'N5';
      },
    }),
  );

  assert.deepEqual(result, { text: 'これはテスト', tokens: null });
  assert.equal(frequencyRequested, false);
  assert.equal(jlptLookupCalls, 0);
  assert.equal(mecabCalls, 0);
});

test('tokenizeSubtitle excludes Yomitan token groups without dictionary headwords from annotation paths', async () => {
  let jlptLookupCalls = 0;
  let frequencyLookupCalls = 0;

  const result = await tokenizeSubtitle(
    '(ダクネスの荒い息) 猫',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [{ text: '(ダクネスの荒い息)', reading: 'だくねすのあらいいき' }],
                    [{ text: '猫', reading: 'ねこ', headwords: [[{ term: '猫' }]] }],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
      getJlptLevel: (text) => {
        jlptLookupCalls += 1;
        return text === '猫' ? 'N5' : null;
      },
      getFrequencyRank: () => {
        frequencyLookupCalls += 1;
        return 12;
      },
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '猫');
  assert.equal(result.tokens?.[0]?.headword, '猫');
  assert.equal(jlptLookupCalls, 1);
  assert.equal(frequencyLookupCalls, 1);
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

test('tokenizeSubtitle uses Yomitan parser result when available and drops no-headword groups', async () => {
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
  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '猫');
  assert.equal(result.tokens?.[0]?.reading, 'ねこ');
  assert.equal(result.tokens?.[0]?.isKnown, false);
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

test('tokenizeSubtitle still assigns frequency to non-known multi-character Yomitan tokens', async () => {
  const result = await tokenizeSubtitle(
    '小園友達',
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
                      text: '友達',
                      reading: 'ともだち',
                      headwords: [[{ term: '友達' }]],
                    },
                  ],
                ],
              },
            ],
          },
        }) as unknown as Electron.BrowserWindow,
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '小園' ? 75 : text === '友達' ? 3000 : null),
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

test('tokenizeSubtitle does not select kana-only N+1 target tokens', async () => {
  const result = await tokenizeSubtitle(
    '私のばあい',
    makeDepsFromYomitanTokens(
      [
        { surface: '私', reading: 'わたし', headword: '私' },
        { surface: 'の', reading: 'の', headword: 'の' },
        { surface: 'ばあい', reading: 'ばあい', headword: '場合' },
      ],
      {
        getMinSentenceWordsForNPlusOne: () => 2,
        isKnownWord: (text) => text === '私',
      },
    ),
  );

  assert.equal(result.tokens?.length, 3);
  assert.equal(
    result.tokens?.some((token) => token.isNPlusOneTarget),
    false,
  );
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
  assert.equal(gaToken?.isKnown, false);
  assert.equal(gaToken?.isNPlusOneTarget, false);
  assert.equal(gaToken?.jlptLevel, undefined);
  assert.equal(gaToken?.frequencyRank, undefined);
  assert.equal(desuToken?.pos1, '助動詞');
  assert.equal(desuToken?.isKnown, false);
  assert.equal(desuToken?.isNPlusOneTarget, false);
  assert.equal(desuToken?.jlptLevel, undefined);
  assert.equal(desuToken?.frequencyRank, undefined);
  assert.equal(targets.length, 1);
  assert.equal(targets[0]?.surface, '仮面');
});

test('tokenizeSubtitle preserves merged token frequency when MeCab positions cross a newline gap', async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async (script: string) => {
        if (script.includes('getTermFrequencies')) {
          return script.includes('"term":"陰に","reading":"いんに"')
            ? [
                {
                  term: '陰に',
                  reading: 'いんに',
                  dictionary: 'JPDBv2㋕',
                  frequency: 5702,
                  displayValue: '5702',
                  displayValueParsed: false,
                },
              ]
            : [];
        }

        return [
          {
            surface: 'X',
            reading: 'えっくす',
            headword: 'X',
            startPos: 0,
            endPos: 1,
          },
          {
            surface: '陰に',
            reading: 'いんに',
            headword: '陰に',
            startPos: 2,
            endPos: 4,
          },
          {
            surface: '潜み',
            reading: 'ひそ',
            headword: '潜む',
            startPos: 4,
            endPos: 6,
          },
        ];
      },
    },
  } as unknown as Electron.BrowserWindow;

  const deps = createTokenizerDepsRuntime({
    getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
    getYomitanParserWindow: () => parserWindow,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    isKnownWord: () => false,
    getKnownWordMatchMode: () => 'headword',
    getJlptLevel: () => null,
    getFrequencyDictionaryEnabled: () => true,
    getMecabTokenizer: () => ({
      tokenize: async () => [
        {
          word: 'X',
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          pos2: '一般',
          pos3: '',
          pos4: '',
          inflectionType: '',
          inflectionForm: '',
          headword: 'X',
          katakanaReading: 'エックス',
          pronunciation: 'エックス',
        },
        {
          word: '陰',
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          pos2: '一般',
          pos3: '',
          pos4: '',
          inflectionType: '',
          inflectionForm: '',
          headword: '陰',
          katakanaReading: 'カゲ',
          pronunciation: 'カゲ',
        },
        {
          word: 'に',
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          pos2: '格助詞',
          pos3: '一般',
          pos4: '',
          inflectionType: '',
          inflectionForm: '',
          headword: 'に',
          katakanaReading: 'ニ',
          pronunciation: 'ニ',
        },
        {
          word: '潜み',
          partOfSpeech: PartOfSpeech.verb,
          pos1: '動詞',
          pos2: '自立',
          pos3: '',
          pos4: '',
          inflectionType: '五段・マ行',
          inflectionForm: '連用形',
          headword: '潜む',
          katakanaReading: 'ヒソミ',
          pronunciation: 'ヒソミ',
        },
      ],
    }),
  });

  const result = await tokenizeSubtitle('X\n陰に潜み', deps);

  assert.equal(result.tokens?.[1]?.surface, '陰に');
  assert.equal(result.tokens?.[1]?.pos1, '名詞|助詞');
  assert.equal(result.tokens?.[1]?.pos2, '一般|格助詞');
  assert.equal(result.tokens?.[1]?.frequencyRank, 5702);
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

test('tokenizeSubtitle uses frequency surface match mode when configured', async () => {
  const result = await tokenizeSubtitle(
    '鍛えた',
    makeDepsFromYomitanTokens([{ surface: '鍛えた', reading: 'きたえた', headword: '鍛える' }], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyDictionaryMatchMode: () => 'surface',
      getFrequencyRank: (text) => (text === '鍛えた' ? 2847 : null),
    }),
  );

  assert.equal(result.text, '鍛えた');
  assert.equal(result.tokens?.[0]?.frequencyRank, 2847);
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

test('createTokenizerDepsRuntime skips known-word lookup for MeCab POS enrichment tokens', async () => {
  let knownWordCalls = 0;

  const deps = createTokenizerDepsRuntime({
    getYomitanExt: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    isKnownWord: () => {
      knownWordCalls += 1;
      return true;
    },
    getKnownWordMatchMode: () => 'headword',
    getJlptLevel: () => null,
    getMecabTokenizer: () => ({
      tokenize: async () => [
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
      ],
    }),
  });

  const tokens = await deps.tokenizeWithMecab('仮面');

  assert.equal(knownWordCalls, 0);
  assert.equal(tokens?.[0]?.isKnown, false);
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

test('tokenizeSubtitle uses Yomitan word classes to classify standalone particles', async () => {
  let mecabCalls = 0;
  const result = await tokenizeSubtitle(
    'は',
    makeDepsFromYomitanTokens(
      [{ surface: 'は', reading: 'は', headword: 'は', wordClasses: ['prt'] }],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === 'は' ? 10 : null),
        getJlptLevel: (text) => (text === 'は' ? 'N5' : null),
        tokenizeWithMecab: async () => {
          mecabCalls += 1;
          return null;
        },
      },
    ),
  );

  assert.equal(mecabCalls, 1);
  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.partOfSpeech, PartOfSpeech.particle);
  assert.equal(result.tokens?.[0]?.pos1, '助詞');
  assert.equal(result.tokens?.[0]?.isNPlusOneTarget, false);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test('tokenizeSubtitle uses Yomitan word classes to classify auxiliary subclasses', async () => {
  const result = await tokenizeSubtitle(
    'です',
    makeDepsFromYomitanTokens(
      [{ surface: 'です', reading: 'です', headword: 'です', wordClasses: ['aux-v'] }],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: () => 10,
        getJlptLevel: () => 'N5',
        tokenizeWithMecab: async () => null,
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.partOfSpeech, PartOfSpeech.bound_auxiliary);
  assert.equal(result.tokens?.[0]?.pos1, '助動詞');
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test('tokenizeSubtitle fills detailed MeCab POS when Yomitan word class supplies coarse POS', async () => {
  const result = await tokenizeSubtitle(
    'は',
    makeDepsFromYomitanTokens(
      [{ surface: 'は', reading: 'は', headword: 'は', wordClasses: ['prt'] }],
      {
        tokenizeWithMecab: async () => [
          {
            headword: 'は',
            surface: 'は',
            reading: 'ハ',
            startPos: 0,
            endPos: 1,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '係助詞',
            pos3: '*',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
      },
    ),
  );

  assert.equal(result.tokens?.[0]?.partOfSpeech, PartOfSpeech.particle);
  assert.equal(result.tokens?.[0]?.pos1, '助詞');
  assert.equal(result.tokens?.[0]?.pos2, '係助詞');
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

test('tokenizeSubtitle excludes default non-independent pos2 from N+1 and frequency annotations', async () => {
  const result = await tokenizeSubtitle(
    'になれば',
    makeDepsFromYomitanTokens([{ surface: 'になれば', reading: 'になれば', headword: 'なる' }], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === 'なる' ? 11 : null),
      tokenizeWithMecab: async () => [
        {
          headword: 'なる',
          surface: 'になれば',
          reading: 'ニナレバ',
          startPos: 0,
          endPos: 4,
          partOfSpeech: PartOfSpeech.verb,
          pos1: '動詞',
          pos2: '非自立',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getMinSentenceWordsForNPlusOne: () => 1,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[0]?.isNPlusOneTarget, false);
});

test('tokenizeSubtitle preserves known-word highlight for exact non-independent kanji noun tokens', async () => {
  const result = await tokenizeSubtitle(
    'その点',
    makeDepsFromYomitanTokens(
      [
        { surface: 'その', reading: 'その', headword: 'その' },
        { surface: '点', reading: 'てん', headword: '点' },
      ],
      {
        isKnownWord: (text) => text === '点' || text === 'てん',
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '点' ? 1384 : null),
        getJlptLevel: (text) => (text === '点' ? 'N3' : null),
        tokenizeWithMecab: async () => [
          {
            headword: 'その',
            surface: 'その',
            reading: 'ソノ',
            startPos: 0,
            endPos: 2,
            partOfSpeech: PartOfSpeech.other,
            pos1: '連体詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '点',
            surface: '点',
            reading: 'テン',
            startPos: 2,
            endPos: 3,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '非自立',
            pos3: '一般',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
      },
    ),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.surface, '点');
  assert.equal(result.tokens?.[1]?.isKnown, true);
  assert.equal(result.tokens?.[1]?.isNPlusOneTarget, false);
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[1]?.jlptLevel, undefined);
});

test('tokenizeSubtitle keeps mecab-tagged interjections tokenized while clearing annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    'ぐはっ',
    makeDepsFromYomitanTokens([{ surface: 'ぐはっ', reading: 'ぐはっ', headword: 'ぐはっ' }], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: () => 17,
      getJlptLevel: () => 'N5',
      tokenizeWithMecab: async () => [
        {
          headword: 'ぐはっ',
          surface: 'ぐはっ',
          reading: 'グハッ',
          startPos: 0,
          endPos: 3,
          partOfSpeech: PartOfSpeech.other,
          pos1: '感動詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
    }),
  );

  assert.equal(result.text, 'ぐはっ');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      reading: token.reading,
      pos1: token.pos1,
      jlptLevel: token.jlptLevel,
      frequencyRank: token.frequencyRank,
      isKnown: token.isKnown,
      isNPlusOneTarget: token.isNPlusOneTarget,
    })),
    [
      {
        surface: 'ぐはっ',
        headword: 'ぐはっ',
        reading: 'ぐはっ',
        pos1: '感動詞',
        jlptLevel: undefined,
        frequencyRank: undefined,
        isKnown: false,
        isNPlusOneTarget: false,
      },
    ],
  );
});

test('tokenizeSubtitle keeps excluded interjections hoverable while clearing only their annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    'ぐはっ 猫',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '猫' ? 11 : 17),
      getJlptLevel: (text) => (text === '猫' ? 'N5' : null),
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [{ text: 'ぐはっ', reading: 'ぐはっ', headwords: [[{ term: 'ぐはっ' }]] }],
                    [{ text: '猫', reading: 'ねこ', headwords: [[{ term: '猫' }]] }],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => [
        {
          headword: 'ぐはっ',
          surface: 'ぐはっ',
          reading: 'グハッ',
          startPos: 0,
          endPos: 3,
          partOfSpeech: PartOfSpeech.other,
          pos1: '感動詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: '猫',
          surface: '猫',
          reading: 'ネコ',
          startPos: 4,
          endPos: 5,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
    }),
  );

  assert.equal(result.text, 'ぐはっ 猫');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      frequencyRank: token.frequencyRank,
      jlptLevel: token.jlptLevel,
    })),
    [
      { surface: 'ぐはっ', headword: 'ぐはっ', frequencyRank: undefined, jlptLevel: undefined },
      { surface: '猫', headword: '猫', frequencyRank: 11, jlptLevel: 'N5' },
    ],
  );
});

test('tokenizeSubtitle keeps explanatory ending variants hoverable while clearing only their annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    '猫んです',
    makeDepsFromYomitanTokens(
      [
        { surface: '猫', reading: 'ねこ', headword: '猫' },
        { surface: 'んです', reading: 'んです', headword: 'ん' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '猫' ? 11 : 500),
        getJlptLevel: (text) => (text === '猫' ? 'N5' : null),
        tokenizeWithMecab: async () => [
          {
            headword: '猫',
            surface: '猫',
            reading: 'ネコ',
            startPos: 0,
            endPos: 1,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '一般',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'ん',
            surface: 'ん',
            reading: 'ン',
            startPos: 1,
            endPos: 2,
            partOfSpeech: PartOfSpeech.other,
            pos1: '名詞',
            pos2: '非自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'です',
            surface: 'です',
            reading: 'デス',
            startPos: 2,
            endPos: 4,
            partOfSpeech: PartOfSpeech.bound_auxiliary,
            pos1: '助動詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
      },
    ),
  );

  assert.equal(result.text, '猫んです');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      jlptLevel: token.jlptLevel,
      frequencyRank: token.frequencyRank,
    })),
    [
      { surface: '猫', headword: '猫', jlptLevel: 'N5', frequencyRank: 11 },
      { surface: 'んです', headword: 'ん', jlptLevel: undefined, frequencyRank: undefined },
    ],
  );
});

test('tokenizeSubtitle keeps standalone grammar-only tokens hoverable while clearing only their annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    '私はこの猫です',
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '私' ? 50 : text === '猫' ? 11 : 500),
      getJlptLevel: (text) => (text === '私' ? 'N5' : text === '猫' ? 'N5' : null),
      getYomitanExt: () => ({ id: 'dummy-ext' }) as any,
      getYomitanParserWindow: () =>
        ({
          isDestroyed: () => false,
          webContents: {
            executeJavaScript: async (script: string) => {
              if (script.includes('getTermFrequencies')) {
                return [];
              }

              return [
                {
                  source: 'scanning-parser',
                  index: 0,
                  content: [
                    [{ text: '私', reading: 'わたし', headwords: [[{ term: '私' }]] }],
                    [{ text: 'は', reading: 'は', headwords: [[{ term: 'は' }]] }],
                    [{ text: 'この', reading: 'この', headwords: [[{ term: 'この' }]] }],
                    [{ text: '猫', reading: 'ねこ', headwords: [[{ term: '猫' }]] }],
                    [{ text: 'です', reading: 'です', headwords: [[{ term: 'です' }]] }],
                  ],
                },
              ];
            },
          },
        }) as unknown as Electron.BrowserWindow,
      tokenizeWithMecab: async () => [
        {
          headword: '私',
          surface: '私',
          reading: 'ワタシ',
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          pos2: '代名詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: 'は',
          surface: 'は',
          reading: 'ハ',
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          pos2: '係助詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: 'この',
          surface: 'この',
          reading: 'コノ',
          startPos: 2,
          endPos: 4,
          partOfSpeech: PartOfSpeech.other,
          pos1: '連体詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: '猫',
          surface: '猫',
          reading: 'ネコ',
          startPos: 4,
          endPos: 5,
          partOfSpeech: PartOfSpeech.noun,
          pos1: '名詞',
          pos2: '一般',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: 'です',
          surface: 'です',
          reading: 'デス',
          startPos: 5,
          endPos: 7,
          partOfSpeech: PartOfSpeech.bound_auxiliary,
          pos1: '助動詞',
          isMerged: true,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
    }),
  );

  assert.equal(result.text, '私はこの猫です');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      frequencyRank: token.frequencyRank,
      jlptLevel: token.jlptLevel,
    })),
    [
      { surface: '私', headword: '私', frequencyRank: 50, jlptLevel: 'N5' },
      { surface: 'は', headword: 'は', frequencyRank: undefined, jlptLevel: undefined },
      { surface: 'この', headword: 'この', frequencyRank: undefined, jlptLevel: undefined },
      { surface: '猫', headword: '猫', frequencyRank: 11, jlptLevel: 'N5' },
      { surface: 'です', headword: 'です', frequencyRank: undefined, jlptLevel: undefined },
    ],
  );
});

test('tokenizeSubtitle keeps trailing quote-particle merged tokens hoverable while clearing only their annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    'どうしてもって',
    makeDepsFromYomitanTokens(
      [{ surface: 'どうしてもって', reading: 'どうしてもって', headword: 'どうしても' }],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === 'どうしても' ? 123 : null),
        getJlptLevel: (text) => (text === 'どうしても' ? 'N3' : null),
        tokenizeWithMecab: async () => [
          {
            headword: 'どうしても',
            surface: 'どうしても',
            reading: 'ドウシテモ',
            startPos: 0,
            endPos: 5,
            partOfSpeech: PartOfSpeech.other,
            pos1: '副詞',
            pos2: '一般',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'って',
            surface: 'って',
            reading: 'ッテ',
            startPos: 5,
            endPos: 7,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '格助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
        getMinSentenceWordsForNPlusOne: () => 1,
      },
    ),
  );

  assert.equal(result.text, 'どうしてもって');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      jlptLevel: token.jlptLevel,
      frequencyRank: token.frequencyRank,
    })),
    [
      {
        surface: 'どうしてもって',
        headword: 'どうしても',
        jlptLevel: undefined,
        frequencyRank: undefined,
      },
    ],
  );
});

test('tokenizeSubtitle keeps auxiliary-stem そうだ grammar tails hoverable while clearing annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    '与えるそうだ',
    makeDepsFromYomitanTokens(
      [
        { surface: '与える', reading: 'あたえる', headword: '与える' },
        { surface: 'そうだ', reading: 'そうだ', headword: 'そうだ' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '与える' ? 100 : text === 'そうだ' ? 12 : null),
        getJlptLevel: (text) => (text === '与える' ? 'N3' : text === 'そうだ' ? 'N5' : null),
        tokenizeWithMecab: async () => [
          {
            headword: '与える',
            surface: '与える',
            reading: 'アタエル',
            startPos: 0,
            endPos: 3,
            partOfSpeech: PartOfSpeech.verb,
            pos1: '動詞',
            pos2: '自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'そう',
            surface: 'そう',
            reading: 'ソウ',
            startPos: 3,
            endPos: 5,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '特殊',
            pos3: '助動詞語幹',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'だ',
            surface: 'だ',
            reading: 'ダ',
            startPos: 5,
            endPos: 6,
            partOfSpeech: PartOfSpeech.bound_auxiliary,
            pos1: '助動詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
        getMinSentenceWordsForNPlusOne: () => 1,
      },
    ),
  );

  assert.equal(result.text, '与えるそうだ');
  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      frequencyRank: token.frequencyRank,
      jlptLevel: token.jlptLevel,
    })),
    [
      { surface: '与える', headword: '与える', frequencyRank: 100, jlptLevel: 'N3' },
      { surface: 'そうだ', headword: 'そうだ', frequencyRank: undefined, jlptLevel: undefined },
    ],
  );
});

test('tokenizeSubtitle excludes single-kana merged tokens from frequency highlighting', async () => {
  const result = await tokenizeSubtitle(
    'た',
    makeDepsFromYomitanTokens([{ surface: 'た', reading: 'た', headword: 'た' }], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === 'た' ? 17 : null),
      getMinSentenceWordsForNPlusOne: () => 1,
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle excludes merged kana-only function/content token from frequency and N+1', async () => {
  const result = await tokenizeSubtitle(
    'になれば',
    makeDepsFromYomitanTokens([{ surface: 'になれば', reading: 'になれば', headword: 'なる' }], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === 'なる' ? 13 : null),
      tokenizeWithMecab: async () => [
        {
          headword: 'に',
          surface: 'に',
          reading: 'ニ',
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          pos2: '格助詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: 'なる',
          surface: 'なれ',
          reading: 'ナレ',
          startPos: 1,
          endPos: 3,
          partOfSpeech: PartOfSpeech.verb,
          pos1: '動詞',
          pos2: '自立',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: 'ば',
          surface: 'ば',
          reading: 'バ',
          startPos: 3,
          endPos: 4,
          partOfSpeech: PartOfSpeech.particle,
          pos1: '助詞',
          pos2: '接続助詞',
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getMinSentenceWordsForNPlusOne: () => 1,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.pos1, '助詞|動詞');
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[0]?.isNPlusOneTarget, false);
});

test('tokenizeSubtitle clears all annotations for kana-only demonstrative helper merges', async () => {
  const result = await tokenizeSubtitle(
    'これで実力どおりか',
    makeDepsFromYomitanTokens(
      [
        { surface: 'これで', reading: 'これで', headword: 'これ' },
        { surface: '実力どおり', reading: 'じつりょくどおり', headword: '実力どおり' },
        { surface: 'か', reading: 'か', headword: 'か' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) =>
          text === 'これ' ? 9 : text === '実力どおり' ? 2500 : text === 'か' ? 800 : null,
        getJlptLevel: (text) =>
          text === 'これ' ? 'N5' : text === '実力どおり' ? 'N1' : text === 'か' ? 'N5' : null,
        isKnownWord: (text) => text === 'これ',
        getMinSentenceWordsForNPlusOne: () => 1,
        tokenizeWithMecab: async () => [
          {
            headword: 'これ',
            surface: 'これ',
            reading: 'コレ',
            startPos: 0,
            endPos: 2,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '代名詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'で',
            surface: 'で',
            reading: 'デ',
            startPos: 2,
            endPos: 3,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '格助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '実力どおり',
            surface: '実力どおり',
            reading: 'ジツリョクドオリ',
            startPos: 3,
            endPos: 8,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '一般',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'か',
            surface: 'か',
            reading: 'カ',
            startPos: 8,
            endPos: 9,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '終助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
      },
    ),
  );

  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      isKnown: token.isKnown,
      isNPlusOneTarget: token.isNPlusOneTarget,
      frequencyRank: token.frequencyRank,
      jlptLevel: token.jlptLevel,
    })),
    [
      {
        surface: 'これで',
        headword: 'これ',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: undefined,
        jlptLevel: undefined,
      },
      {
        surface: '実力どおり',
        headword: '実力どおり',
        isKnown: false,
        isNPlusOneTarget: true,
        frequencyRank: 2500,
        jlptLevel: 'N1',
      },
      {
        surface: 'か',
        headword: 'か',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: undefined,
        jlptLevel: undefined,
      },
    ],
  );
});

test('tokenizeSubtitle clears all annotations for explanatory pondering endings', async () => {
  const result = await tokenizeSubtitle(
    '俺どうかしちゃったのかな',
    makeDepsFromYomitanTokens(
      [
        { surface: '俺', reading: 'おれ', headword: '俺' },
        { surface: 'どうかしちゃった', reading: 'どうかしちゃった', headword: 'どうかしちゃう' },
        { surface: 'のかな', reading: 'のかな', headword: 'の' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '俺' ? 19 : text === 'どうかしちゃう' ? 3200 : 77),
        getJlptLevel: (text) =>
          text === '俺' ? 'N5' : text === 'どうかしちゃう' ? 'N3' : text === 'の' ? 'N5' : null,
        isKnownWord: (text) => text === '俺' || text === 'の',
        getMinSentenceWordsForNPlusOne: () => 1,
        tokenizeWithMecab: async () => [
          {
            headword: '俺',
            surface: '俺',
            reading: 'オレ',
            startPos: 0,
            endPos: 1,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '代名詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'どうかしちゃう',
            surface: 'どうかしちゃった',
            reading: 'ドウカシチャッタ',
            startPos: 1,
            endPos: 8,
            partOfSpeech: PartOfSpeech.verb,
            pos1: '動詞',
            pos2: '自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'の',
            surface: 'のかな',
            reading: 'ノカナ',
            startPos: 8,
            endPos: 11,
            partOfSpeech: PartOfSpeech.other,
            pos1: '名詞|助動詞',
            pos2: '非自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
      },
    ),
  );

  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      isKnown: token.isKnown,
      isNPlusOneTarget: token.isNPlusOneTarget,
      frequencyRank: token.frequencyRank,
      jlptLevel: token.jlptLevel,
    })),
    [
      {
        surface: '俺',
        headword: '俺',
        isKnown: true,
        isNPlusOneTarget: false,
        frequencyRank: 19,
        jlptLevel: 'N5',
      },
      {
        surface: 'どうかしちゃった',
        headword: 'どうかしちゃう',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: 3200,
        jlptLevel: 'N3',
      },
      {
        surface: 'のかな',
        headword: 'の',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: undefined,
        jlptLevel: undefined,
      },
    ],
  );
});

test('tokenizeSubtitle keeps frequency for content-led merged token with trailing colloquial suffixes', async () => {
  const result = await tokenizeSubtitle(
    '張り切ってんじゃ',
    makeDepsFromYomitanTokens(
      [{ surface: '張り切ってん', reading: 'はき', headword: '張り切る' }],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '張り切る' ? 5468 : null),
        tokenizeWithMecab: async () => [
          {
            headword: '張り切る',
            surface: '張り切っ',
            reading: 'ハリキッ',
            startPos: 0,
            endPos: 4,
            partOfSpeech: PartOfSpeech.verb,
            pos1: '動詞',
            pos2: '自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'て',
            surface: 'て',
            reading: 'テ',
            startPos: 4,
            endPos: 5,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '接続助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'んじゃ',
            surface: 'んじゃ',
            reading: 'ンジャ',
            startPos: 5,
            endPos: 8,
            partOfSpeech: PartOfSpeech.other,
            pos1: '接続詞',
            pos2: '*',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
        getMinSentenceWordsForNPlusOne: () => 1,
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '張り切ってん');
  assert.equal(result.tokens?.[0]?.pos1, '動詞|助詞|接続詞');
  assert.equal(result.tokens?.[0]?.frequencyRank, 5468);
});

test('tokenizeSubtitle clears all annotations for explanatory contrast endings', async () => {
  const result = await tokenizeSubtitle(
    '最近辛いものが続いとるんですけど',
    makeDepsFromYomitanTokens(
      [
        { surface: '最近', reading: 'さいきん', headword: '最近' },
        { surface: '辛い', reading: 'つらい', headword: '辛い' },
        { surface: 'もの', reading: 'もの', headword: 'もの' },
        { surface: 'が', reading: 'が', headword: 'が' },
        { surface: '続いとる', reading: 'つづいとる', headword: '続く' },
        { surface: 'んですけど', reading: 'んですけど', headword: 'ん' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) =>
          text === '最近' ? 120 : text === '辛い' ? 800 : text === '続く' ? 240 : 77,
        getJlptLevel: (text) =>
          text === '最近' ? 'N4' : text === '辛い' ? 'N2' : text === '続く' ? 'N4' : null,
        isKnownWord: (text) => text === '最近',
        getMinSentenceWordsForNPlusOne: () => 1,
        tokenizeWithMecab: async () => [
          {
            headword: '最近',
            surface: '最近',
            reading: 'サイキン',
            startPos: 0,
            endPos: 2,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '副詞可能',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '辛い',
            surface: '辛い',
            reading: 'ツライ',
            startPos: 2,
            endPos: 4,
            partOfSpeech: PartOfSpeech.i_adjective,
            pos1: '形容詞',
            pos2: '自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'もの',
            surface: 'もの',
            reading: 'モノ',
            startPos: 4,
            endPos: 6,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '一般',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'が',
            surface: 'が',
            reading: 'ガ',
            startPos: 6,
            endPos: 7,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '格助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '続く',
            surface: '続いとる',
            reading: 'ツヅイトル',
            startPos: 7,
            endPos: 11,
            partOfSpeech: PartOfSpeech.verb,
            pos1: '動詞',
            pos2: '自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'ん',
            surface: 'んですけど',
            reading: 'ンデスケド',
            startPos: 11,
            endPos: 16,
            partOfSpeech: PartOfSpeech.other,
            pos1: '名詞|助動詞|助詞',
            pos2: '非自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
      },
    ),
  );

  assert.deepEqual(
    result.tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      isKnown: token.isKnown,
      isNPlusOneTarget: token.isNPlusOneTarget,
      frequencyRank: token.frequencyRank,
      jlptLevel: token.jlptLevel,
    })),
    [
      {
        surface: '最近',
        headword: '最近',
        isKnown: true,
        isNPlusOneTarget: false,
        frequencyRank: 120,
        jlptLevel: 'N4',
      },
      {
        surface: '辛い',
        headword: '辛い',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: 800,
        jlptLevel: 'N2',
      },
      {
        surface: 'もの',
        headword: 'もの',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: 77,
        jlptLevel: undefined,
      },
      {
        surface: 'が',
        headword: 'が',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: undefined,
        jlptLevel: undefined,
      },
      {
        surface: '続いとる',
        headword: '続く',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: 240,
        jlptLevel: 'N4',
      },
      {
        surface: 'んですけど',
        headword: 'ん',
        isKnown: false,
        isNPlusOneTarget: false,
        frequencyRank: undefined,
        jlptLevel: undefined,
      },
    ],
  );
});

test('tokenizeSubtitle clears annotations for ことに while preserving lexical N+1 target', async () => {
  const result = await tokenizeSubtitle(
    'さっきの俺と違うことに気付かないのかい？',
    makeDepsFromYomitanTokens(
      [
        { surface: 'さっき', reading: 'さっき', headword: 'さっき' },
        { surface: 'の', reading: 'の', headword: 'の' },
        { surface: '俺', reading: 'おれ', headword: '俺' },
        { surface: 'と', reading: 'と', headword: 'と' },
        { surface: '違う', reading: 'ちがう', headword: '違う' },
        { surface: 'ことに', reading: 'ことに', headword: '事' },
        { surface: '気付かない', reading: 'きづかない', headword: '気付く' },
        { surface: 'の', reading: 'の', headword: 'の' },
        { surface: 'かい', reading: 'かい', headword: 'かい' },
        { surface: '？', reading: '', headword: '？' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) =>
          text === '違う' ? 900 : text === '事' ? 81 : text === '気付く' ? 1500 : null,
        getJlptLevel: (text) =>
          text === '違う' ? 'N4' : text === '事' ? 'N4' : text === '気付く' ? 'N3' : null,
        isKnownWord: (text) => ['さっき', 'の', '俺', 'と', '気付く', 'かい', '？'].includes(text),
        getMinSentenceWordsForNPlusOne: () => 1,
        tokenizeWithMecab: async () => [
          {
            headword: 'さっき',
            surface: 'さっき',
            reading: 'サッキ',
            startPos: 0,
            endPos: 3,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '副詞可能',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'の',
            surface: 'の',
            reading: 'ノ',
            startPos: 3,
            endPos: 4,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '連体化',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '俺',
            surface: '俺',
            reading: 'オレ',
            startPos: 4,
            endPos: 5,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '代名詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'と',
            surface: 'と',
            reading: 'ト',
            startPos: 5,
            endPos: 6,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '格助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '違う',
            surface: '違う',
            reading: 'チガウ',
            startPos: 6,
            endPos: 8,
            partOfSpeech: PartOfSpeech.verb,
            pos1: '動詞',
            pos2: '自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '事',
            surface: 'こと',
            reading: 'コト',
            startPos: 8,
            endPos: 10,
            partOfSpeech: PartOfSpeech.noun,
            pos1: '名詞',
            pos2: '非自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'に',
            surface: 'に',
            reading: 'ニ',
            startPos: 10,
            endPos: 11,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '格助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '気付く',
            surface: '気付か',
            reading: 'キヅカ',
            startPos: 11,
            endPos: 14,
            partOfSpeech: PartOfSpeech.verb,
            pos1: '動詞',
            pos2: '自立',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'ない',
            surface: 'ない',
            reading: 'ナイ',
            startPos: 14,
            endPos: 16,
            partOfSpeech: PartOfSpeech.bound_auxiliary,
            pos1: '助動詞',
            pos2: '*',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'の',
            surface: 'の',
            reading: 'ノ',
            startPos: 16,
            endPos: 17,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '終助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: 'かい',
            surface: 'かい',
            reading: 'カイ',
            startPos: 17,
            endPos: 19,
            partOfSpeech: PartOfSpeech.particle,
            pos1: '助詞',
            pos2: '終助詞',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
          {
            headword: '？',
            surface: '？',
            reading: '',
            startPos: 19,
            endPos: 20,
            partOfSpeech: PartOfSpeech.symbol,
            pos1: '記号',
            pos2: '一般',
            isMerged: false,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ],
      },
    ),
  );

  const tokenSummary = result.tokens?.map((token) => ({
    surface: token.surface,
    headword: token.headword,
    isKnown: token.isKnown,
    isNPlusOneTarget: token.isNPlusOneTarget,
    frequencyRank: token.frequencyRank,
    jlptLevel: token.jlptLevel,
  }));

  assert.deepEqual(
    tokenSummary?.find((token) => token.surface === 'ことに'),
    {
      surface: 'ことに',
      headword: '事',
      isKnown: false,
      isNPlusOneTarget: false,
      frequencyRank: undefined,
      jlptLevel: undefined,
    },
  );
  assert.deepEqual(
    tokenSummary?.find((token) => token.surface === '違う'),
    {
      surface: '違う',
      headword: '違う',
      isKnown: false,
      isNPlusOneTarget: true,
      frequencyRank: 900,
      jlptLevel: 'N4',
    },
  );
});

test('tokenizeSubtitle excludes default non-independent pos2 from N+1 when JLPT/frequency are disabled', async () => {
  let mecabCalls = 0;
  const result = await tokenizeSubtitle(
    'になれば',
    makeDepsFromYomitanTokens([{ surface: 'になれば', reading: 'になれば', headword: 'なる' }], {
      getJlptEnabled: () => false,
      getFrequencyDictionaryEnabled: () => false,
      getMinSentenceWordsForNPlusOne: () => 1,
      tokenizeWithMecab: async () => {
        mecabCalls += 1;
        return [
          {
            headword: 'なる',
            surface: 'になれば',
            reading: 'ニナレバ',
            startPos: 0,
            endPos: 4,
            partOfSpeech: PartOfSpeech.verb,
            pos1: '動詞',
            pos2: '非自立',
            isMerged: true,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ];
      },
    }),
  );

  assert.equal(mecabCalls, 1);
  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.isNPlusOneTarget, false);
});
