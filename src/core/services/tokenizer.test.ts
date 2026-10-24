import test from 'node:test';
import assert from 'node:assert/strict';
import { MergedToken, PartOfSpeech, Token } from '../../types';
import {
  createTokenizerDepsRuntime,
  TokenizerDepsRuntimeOptions,
  TokenizerServiceDeps,
  tokenizeSubtitle,
} from './tokenizer';
import { createBackendDeps, runInjectedYomitanScript } from './tokenizer/yomitan-scan-test-harness';

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

type TermReadingPair = { term: string; reading: string | null };

interface YomitanBackend {
  // Answers the scanner call and the parseText fallback.
  respond: (script: string) => unknown;
  // Answers getTermFrequencies for the requested term/reading pairs.
  frequencies?: (pairs: TermReadingPair[]) => unknown[] | Promise<unknown[]>;
  // Enabled dictionaries of the active profile, in priority order.
  dictionaries?: string[];
  dictionaryInfo?: unknown[];
}

// A ready Yomitan parser window. Profile-metadata and frequency scripts run
// for real against a fake backend; every other script goes to `respond`.
function createYomitanParserWindow(backend: YomitanBackend): Electron.BrowserWindow {
  const dictionaries = backend.dictionaries ?? ['freq-dict'];
  const handleBackendAction = (action: string, params: unknown): unknown => {
    switch (action) {
      case 'optionsGetFull':
        return {
          profileCurrent: 0,
          profiles: [
            {
              options: {
                scanning: { length: 40 },
                dictionaries: dictionaries.map((name, id) => ({ name, enabled: true, id })),
              },
            },
          ],
        };
      case 'getDictionaryInfo':
        return backend.dictionaryInfo ?? [];
      case 'getTermFrequencies':
        return (
          backend.frequencies?.(
            (params as { termReadingList: TermReadingPair[] }).termReadingList,
          ) ?? []
        );
      default:
        throw new Error(`unexpected action: ${action}`);
    }
  };
  return {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async (script: string) =>
        script.includes('getDictionaryInfo') || script.includes('getTermFrequencies')
          ? await runInjectedYomitanScript(script, handleBackendAction)
          : backend.respond(script),
    },
  } as unknown as Electron.BrowserWindow;
}

function makeYomitanDeps(
  backend: YomitanBackend,
  overrides: Partial<TokenizerServiceDeps> = {},
): TokenizerServiceDeps {
  const parserWindow = createYomitanParserWindow(backend);
  return makeDeps({
    getYomitanExt: () => ({ id: 'dummy-ext' }) as Electron.Extension,
    getYomitanParserWindow: () => parserWindow,
    ...overrides,
  });
}

interface YomitanTokenInput {
  surface: string;
  reading?: string;
  headword?: string;
  // Defaults to the end of the previous token.
  startPos?: number;
  frequencyRank?: number;
  isNameMatch?: boolean;
  wordClasses?: string[];
  isUnparsedRun?: boolean;
}

// Scanner-call responder returning `tokens` as the in-window scanner would.
function scanTokens(tokens: YomitanTokenInput[]) {
  return () => {
    let cursor = 0;
    return tokens.map((token) => {
      const startPos = token.startPos ?? cursor;
      const endPos = startPos + token.surface.length;
      cursor = endPos;
      return {
        surface: token.surface,
        reading: token.reading ?? token.surface,
        headword: token.headword ?? token.surface,
        startPos,
        endPos,
        isNameMatch: token.isNameMatch ?? false,
        frequencyRank: token.frequencyRank,
        wordClasses: token.wordClasses,
        isUnparsedRun: token.isUnparsedRun,
      };
    });
  };
}

function makeDepsFromYomitanTokens(
  tokens: YomitanTokenInput[],
  overrides: Partial<TokenizerServiceDeps> = {},
  backend: Omit<YomitanBackend, 'respond'> = {},
): TokenizerServiceDeps {
  return makeYomitanDeps({ ...backend, respond: scanTokens(tokens) }, overrides);
}

interface ParseSegment {
  text: string;
  reading: string;
  headwords?: string[];
}

// One parseText segment; each headword is one alternative dictionary term.
function seg(text: string, reading: string, ...headwords: string[]): ParseSegment {
  return headwords.length > 0 ? { text, reading, headwords } : { text, reading };
}

// Deps whose scanner yields no usable payload, so tokenization falls back to
// parseText and gets `lines` (token groups of segments) as the single
// scanning-parser candidate.
function makeDepsFromScanningParser(
  lines: ParseSegment[][],
  overrides: Partial<TokenizerServiceDeps> = {},
  backend: Omit<YomitanBackend, 'respond'> = {},
): TokenizerServiceDeps {
  const parseResults = [
    {
      source: 'scanning-parser',
      index: 0,
      content: lines.map((line) =>
        line.map(({ text, reading, headwords }) => ({
          text,
          reading,
          ...(headwords ? { headwords: headwords.map((term) => [{ term }]) } : {}),
        })),
      ),
    },
  ];
  return makeYomitanDeps({ ...backend, respond: () => parseResults }, overrides);
}

function yomitanFrequency(
  term: string,
  reading: string | null,
  frequency: number,
  extra: Record<string, unknown> = {},
) {
  return {
    term,
    reading,
    dictionary: 'freq-dict',
    frequency,
    displayValue: String(frequency),
    displayValueParsed: true,
    ...extra,
  };
}

function requestsPair(pairs: TermReadingPair[], term: string, reading: string | null): boolean {
  return pairs.some((pair) => pair.term === term && pair.reading === reading);
}

// A MeCab token for POS enrichment, which only reads its span and POS tags.
function mecabToken(
  surface: string,
  startPos: number,
  pos1: string,
  pos2?: string,
  pos3?: string,
): MergedToken {
  return {
    surface,
    headword: surface,
    reading: '',
    startPos,
    endPos: startPos + surface.length,
    partOfSpeech: PartOfSpeech.other,
    pos1,
    pos2,
    pos3,
    isMerged: false,
    isKnown: false,
    isNPlusOneTarget: false,
  };
}

// A raw MeCab tokenizer word, before createTokenizerDepsRuntime merges it.
function mecabWord(
  word: string,
  partOfSpeech: PartOfSpeech,
  pos1: string,
  pos2: string,
  katakanaReading: string,
  extra: Partial<Token> = {},
): Token {
  return {
    word,
    partOfSpeech,
    pos1,
    pos2,
    pos3: '',
    pos4: '',
    inflectionType: '',
    inflectionForm: '',
    headword: word,
    katakanaReading,
    pronunciation: katakanaReading,
    ...extra,
  };
}

function makeRuntimeOptions(
  overrides: Partial<TokenizerDepsRuntimeOptions> = {},
): TokenizerDepsRuntimeOptions {
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
    getMecabTokenizer: () => null,
    ...overrides,
  };
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

test('tokenizeSubtitle splits same-line grammar endings before applying annotations', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromScanningParser([[seg('猫', 'ねこ', '猫'), seg('です', 'です', 'です')]], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '猫' ? 40 : text === 'です' ? 50 : null),
      getJlptLevel: (text) => (text === '猫' || text === 'です' ? 'N5' : null),
      isKnownWord: (text) => text === 'です',
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, '猫');
  assert.equal(result.tokens?.[0]?.jlptLevel, 'N5');
  assert.equal(result.tokens?.[0]?.frequencyRank, 40);
  assert.equal(result.tokens?.[1]?.surface, 'です');
  assert.equal(result.tokens?.[1]?.isKnown, true);
  assert.equal(result.tokens?.[1]?.isNPlusOneTarget, false);
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[1]?.jlptLevel, undefined);
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
  assert.equal(result.tokens?.[0]?.isNameMatch, true);
  assert.equal(result.tokens?.[1]?.isNameMatch, false);
});

test('tokenizeSubtitle attaches character image metadata to name matches when enabled', async () => {
  const result = await tokenizeSubtitle(
    'アクアです',
    makeDepsFromYomitanTokens(
      [
        { surface: 'アクア', reading: 'あくあ', headword: 'アクア', isNameMatch: true },
        { surface: 'です', reading: 'です', headword: 'です' },
      ],
      {
        getNameMatchImagesEnabled: () => true,
        getCharacterNameImage: (term) =>
          term === 'アクア'
            ? {
                src: 'data:image/png;base64,AAAA',
                alt: 'アクア',
              }
            : null,
      } as Partial<TokenizerServiceDeps>,
    ),
  );

  assert.deepEqual(result.tokens?.[0]?.characterImage, {
    src: 'data:image/png;base64,AAAA',
    alt: 'アクア',
  });
  assert.equal(result.tokens?.[1]?.characterImage, undefined);
});

test('tokenizeSubtitle keeps tokens when character image lookup throws', async () => {
  const result = await tokenizeSubtitle(
    'アクア',
    makeDepsFromYomitanTokens(
      [{ surface: 'アクア', reading: 'あくあ', headword: 'アクア', isNameMatch: true }],
      {
        getNameMatchImagesEnabled: () => true,
        getCharacterNameImage: () => {
          throw new Error('image lookup failed');
        },
      } as Partial<TokenizerServiceDeps>,
    ),
  );

  assert.equal(result.tokens?.[0]?.surface, 'アクア');
  assert.equal(result.tokens?.[0]?.characterImage, undefined);
});

test('tokenizeSubtitle omits character image metadata when name-match images are disabled', async () => {
  const result = await tokenizeSubtitle(
    'アクア',
    makeDepsFromYomitanTokens(
      [{ surface: 'アクア', reading: 'あくあ', headword: 'アクア', isNameMatch: true }],
      {
        getNameMatchImagesEnabled: () => false,
        getCharacterNameImage: () => ({
          src: 'data:image/png;base64,AAAA',
          alt: 'アクア',
        }),
      } as Partial<TokenizerServiceDeps>,
    ),
  );

  assert.equal(result.tokens?.[0]?.characterImage, undefined);
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

test('tokenizeSubtitle uses left-to-right yomitan scanning to keep full katakana name tokens', async () => {
  const result = await tokenizeSubtitle(
    'カズマ 魔王軍',
    makeDepsFromYomitanTokens([
      { surface: 'カズマ', reading: 'かずま' },
      { surface: '魔王軍', reading: 'まおうぐん', startPos: 4 },
    ]),
  );

  assert.deepEqual(
    result.tokens?.map(({ surface, reading, headword, startPos, endPos }) => ({
      surface,
      reading,
      headword,
      startPos,
      endPos,
    })),
    [
      { surface: 'カズマ', reading: 'かずま', headword: 'カズマ', startPos: 0, endPos: 3 },
      { surface: '魔王軍', reading: 'まおうぐん', headword: '魔王軍', startPos: 4, endPos: 7 },
    ],
  );
});

test('tokenizeSubtitle loads frequency ranks from Yomitan installed dictionaries', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromScanningParser(
      [[seg('猫', 'ねこ', '猫')]],
      { getFrequencyDictionaryEnabled: () => true },
      { frequencies: () => [yomitanFrequency('猫', 'ねこ', 77)] },
    ),
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
    makeDepsFromScanningParser(
      [[seg('猫', 'ねこ', '猫')]],
      {
        getFrequencyDictionaryEnabled: () => true,
        tokenizeWithMecab: async () => {
          mecabRequested = true;
          return await mecabDeferred.promise;
        },
      },
      {
        frequencies: async () => {
          frequencyRequested = true;
          return await frequencyDeferred.promise;
        },
      },
    ),
  );

  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(frequencyRequested, true);
  assert.equal(mecabRequested, true);

  frequencyDeferred.resolve([yomitanFrequency('猫', 'ねこ', 77)]);
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
    makeDepsFromScanningParser(
      [[seg('猫', 'ねこ', '猫')]],
      {
        onTokenizationReady: (text) => {
          tokenizationReadyText = text;
        },
        getFrequencyDictionaryEnabled: () => true,
        tokenizeWithMecab: async () => await mecabDeferred.promise,
      },
      { frequencies: async () => await frequencyDeferred.promise },
    ),
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
    makeDepsFromScanningParser([
      [seg('断', 'だん', '断じて'), seg('じて', '', 'じて')],
      [seg('見', 'み', '見る'), seg('ていない', '', 'ていない')],
    ]),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, '断じて');
  assert.equal(result.tokens?.[0]?.reading, 'だんじて');
  assert.equal(result.tokens?.[1]?.surface, '見ていない');
  assert.equal(result.tokens?.[1]?.reading, 'み');
});

test('tokenizeSubtitle queries headword frequencies with token reading for disambiguation', async () => {
  const requested: TermReadingPair[] = [];
  const result = await tokenizeSubtitle(
    '鍛えた',
    makeDepsFromScanningParser(
      [[seg('鍛えた', 'きた', '鍛える')]],
      { getFrequencyDictionaryEnabled: () => true },
      {
        frequencies: (pairs) => {
          requested.push(...pairs);
          return requestsPair(pairs, '鍛える', 'きた')
            ? [yomitanFrequency('鍛える', 'きたえる', 46961, { displayValue: '2847,46961' })]
            : [];
        },
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.headword, '鍛える');
  assert.equal(result.tokens?.[0]?.reading, 'きた');
  assert.equal(result.tokens?.[0]?.frequencyRank, 2847);
  // The reading lookup hits, so no term-only fallback pair is requested.
  assert.deepEqual(
    requested.filter((pair) => pair.term === '鍛える'),
    [{ term: '鍛える', reading: 'きた' }],
  );
});

test('tokenizeSubtitle falls back to term-only Yomitan frequency lookup when reading is noisy', async () => {
  const result = await tokenizeSubtitle(
    '断じて',
    makeDepsFromScanningParser(
      [[seg('断じて', 'だん', '断じて')]],
      { getFrequencyDictionaryEnabled: () => true },
      {
        frequencies: (pairs) =>
          requestsPair(pairs, '断じて', null) ? [yomitanFrequency('断じて', null, 7082)] : [],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 7082);
});

test('tokenizeSubtitle avoids headword term-only fallback rank when reading-specific frequency exists', async () => {
  const cc100 = { dictionary: 'CC100', dictionaryPriority: 0, displayValue: null };
  const result = await tokenizeSubtitle(
    '無人',
    makeDepsFromScanningParser(
      [[seg('無人', 'むじん', '無人')]],
      { getFrequencyDictionaryEnabled: () => true },
      {
        frequencies: (pairs) =>
          requestsPair(pairs, '無人', 'むじん')
            ? [
                yomitanFrequency('無人', null, 157632, cc100),
                yomitanFrequency('無人', 'むじん', 7141, cc100),
              ]
            : [],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 7141);
});

test('tokenizeSubtitle prefers Yomitan frequency from highest-priority dictionary', async () => {
  const result = await tokenizeSubtitle(
    '猫',
    makeDepsFromScanningParser(
      [[seg('猫', 'ねこ', '猫')]],
      { getFrequencyDictionaryEnabled: () => true },
      {
        frequencies: () => [
          yomitanFrequency('猫', 'ねこ', 5, { dictionary: 'low-priority', dictionaryPriority: 2 }),
          yomitanFrequency('猫', 'ねこ', 100, {
            dictionary: 'high-priority',
            dictionaryPriority: 0,
          }),
        ],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 100);
});

// 潜み scanned with a truncated reading and no scan-derived rank, so the rank
// comes from the term-only CC100 lookup.
const HISOMI_TOKEN: YomitanTokenInput = { surface: '潜み', reading: 'ひそ', headword: '潜む' };
const CC100_HISOMI = {
  dictionary: 'CC100',
  hasReading: false,
  displayValue: null,
  displayValueParsed: false,
};

test('tokenizeSubtitle ignores occurrence-based Yomitan frequencies for inflected terms', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDepsFromYomitanTokens(
      [HISOMI_TOKEN],
      { getFrequencyDictionaryEnabled: () => true },
      {
        dictionaries: ['CC100'],
        dictionaryInfo: [{ title: 'CC100', frequencyMode: 'occurrence-based' }],
        frequencies: () => [yomitanFrequency('潜む', 'ひそ', 118121, CC100_HISOMI)],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle falls back to raw term-only Yomitan rank when no scan-derived rank exists', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDepsFromYomitanTokens(
      [HISOMI_TOKEN],
      { getFrequencyDictionaryEnabled: () => true },
      {
        dictionaries: ['CC100'],
        dictionaryInfo: [{ title: 'CC100', frequencyMode: 'rank-based' }],
        frequencies: () => [yomitanFrequency('潜む', 'ひそ', 118121, CC100_HISOMI)],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 118121);
});

test('tokenizeSubtitle keeps parsed display rank for term-only inflected headword fallback', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDepsFromYomitanTokens(
      [HISOMI_TOKEN],
      { getFrequencyDictionaryEnabled: () => true },
      {
        dictionaries: ['CC100'],
        dictionaryInfo: [{ title: 'CC100', frequencyMode: 'rank-based' }],
        frequencies: () => [
          yomitanFrequency('潜む', 'ひそ', 118121, { ...CC100_HISOMI, displayValue: '118,121' }),
        ],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 118);
});

test('tokenizeSubtitle preserves scan-derived rank over lower-priority Yomitan fallback', async () => {
  const result = await tokenizeSubtitle(
    '潜み',
    makeDepsFromYomitanTokens(
      [{ ...HISOMI_TOKEN, reading: 'ひそむ', frequencyRank: 4073 }],
      { getFrequencyDictionaryEnabled: () => true },
      {
        frequencies: () => [
          yomitanFrequency('潜む', 'ひそ', 118121, { ...CC100_HISOMI, dictionaryPriority: 2 }),
        ],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 4073);
});

test('tokenizeSubtitle uses only selected Yomitan headword for frequency lookup', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromScanningParser([[seg('猫です', 'ねこです', '猫です', '猫')]], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '猫' ? 40 : text === '猫です' ? 1200 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 1200);
});

test('tokenizeSubtitle keeps furigana-split Yomitan segments as one token', async () => {
  const result = await tokenizeSubtitle(
    '友達と話した',
    makeDepsFromScanningParser(
      [
        [seg('友', 'とも', '友達'), seg('達', 'だち')],
        [seg('と', 'と', 'と')],
        [seg('話した', 'はなした', '話す')],
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '友達' ? 22 : text === '話す' ? 90 : null),
      },
    ),
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
    makeDepsFromScanningParser([[seg('猫', 'ねこ', 'ネコ')]], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '猫' ? 1200 : text === 'ネコ' ? 8 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 8);
});

test('tokenizeSubtitle falls back to exact surface frequency when merged headword lookup misses', async () => {
  const requested: TermReadingPair[] = [];
  const result = await tokenizeSubtitle(
    '陰に',
    makeDepsFromScanningParser(
      [[seg('陰に', 'いんに', '陰')]],
      { getFrequencyDictionaryEnabled: () => true },
      {
        frequencies: (pairs) => {
          requested.push(...pairs);
          return requestsPair(pairs, '陰に', 'いんに')
            ? [yomitanFrequency('陰に', 'いんに', 5702)]
            : [];
        },
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '陰に');
  assert.equal(result.tokens?.[0]?.headword, '陰');
  assert.equal(result.tokens?.[0]?.frequencyRank, 5702);
  assert.equal(requestsPair(requested, '陰', 'いんに'), true);
  assert.equal(requestsPair(requested, '陰に', 'いんに'), true);
});

test('tokenizeSubtitle keeps no frequency when only reading matches and headword misses', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromScanningParser([[seg('猫', 'ねこ', '猫です')]], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === 'ねこ' ? 77 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle ignores invalid frequency rank on selected headword', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromScanningParser([[seg('猫です', 'ねこです', '猫', '猫です')]], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '猫' ? Number.NaN : text === '猫です' ? 500 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle ignores candidates with no dictionary rank when higher-frequency candidate exists', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromScanningParser([[seg('猫', 'ねこ', '猫', '猫です', 'unknown-term')]], {
      getFrequencyDictionaryEnabled: () => true,
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
    makeDepsFromYomitanTokens([{ surface: '猫', reading: 'ねこ' }], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: () => {
        throw new Error('frequency lookup unavailable');
      },
    }),
  );

  assert.equal(result.tokens?.[0]?.surface, '猫');
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test('tokenizeSubtitle keeps standalone particle token hoverable while clearing annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    'は',
    makeDepsFromScanningParser([[seg('は', 'は', 'は')]], {
      getFrequencyDictionaryEnabled: () => true,
      tokenizeWithMecab: async () => [mecabToken('は', 0, '助詞')],
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

test('tokenizeSubtitle keeps repeated kana interjections tokenized while clearing annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    'ああ',
    makeDepsFromScanningParser([[seg('ああ', 'ああ', 'ああ')]], {
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

test('tokenizeSubtitle returns the normalized text when it comes out empty', async () => {
  // Handing back the original would push whatever normalization dropped into app state
  // as if it were subtitle text.
  const result = await tokenizeSubtitle(' \\n  ', makeDeps());
  assert.deepEqual(result, { text: '', tokens: null });
});

const LINE_NORMALIZATION_CASES = [
  {
    name: 'normalizes newlines before Yomitan parse request',
    input: '猫\\Nです\nね',
    scanned: '猫 です ね',
    text: '猫\nです\nね',
  },
  {
    name: 'collapses zero-width separators before Yomitan parse request',
    input: 'キリキリと​かかってこい\nこのヘナチョコ冒険者どもめが！',
    scanned: 'キリキリと かかってこい このヘナチョコ冒険者どもめが！',
    text: 'キリキリと​かかってこい\nこのヘナチョコ冒険者どもめが！',
  },
];

for (const c of LINE_NORMALIZATION_CASES) {
  test(`tokenizeSubtitle ${c.name}`, async () => {
    const lookups: string[] = [];
    const result = await tokenizeSubtitle(
      c.input,
      makeDeps(createBackendDeps({ dictionaries: ['JMdict'], lookups, termsFind: () => null })),
    );

    // The scanner's first lookup window is the whole normalized line.
    assert.equal(lookups[0], c.scanned);
    assert.equal(result.text, c.text);
    assert.equal(result.tokens, null);
  });
}

test('tokenizeSubtitle preserves CRLF boundaries between simultaneous cues', async () => {
  const result = await tokenizeSubtitle('a\r\n\r\nb', makeDeps());

  assert.deepEqual(result, { text: 'a\n\nb', tokens: null });
});

test('tokenizeSubtitle returns null tokens when Yomitan parsing is unavailable', async () => {
  const result = await tokenizeSubtitle('猫です', makeDeps());

  assert.deepEqual(result, { text: '猫です', tokens: null });
});

test('tokenizeSubtitle uses Yomitan parser result and keeps no-headword groups as surface tokens', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromScanningParser([[seg('猫', 'ねこ', '猫')], [seg('です', 'です')]]),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, '猫');
  assert.equal(result.tokens?.[0]?.reading, 'ねこ');
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.surface, 'です');
  assert.equal(result.tokens?.[1]?.headword, 'です');
});

test('tokenizeSubtitle preserves segmented Yomitan line as one token', async () => {
  const result = await tokenizeSubtitle(
    '猫です',
    makeDepsFromScanningParser([[seg('猫', 'ねこ', '猫です'), seg('です', 'です')]]),
  );

  assert.equal(result.text, '猫です');
  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '猫です');
  assert.equal(result.tokens?.[0]?.reading, 'ねこです');
  assert.equal(result.tokens?.[0]?.headword, '猫です');
  assert.equal(result.tokens?.[0]?.isKnown, false);
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

test('tokenizeSubtitle keeps correct MeCab pos1 enrichment when Yomitan offsets skip spaces', async () => {
  const result = await tokenizeSubtitle(
    '私も あの仮面が欲しいです',
    makeDepsFromScanningParser(
      [
        [seg('私', 'わたし', '私')],
        [seg('も', 'も', 'も')],
        [seg('あの', 'あの', 'あの')],
        [seg('仮面', 'かめん', '仮面')],
        [seg('が', 'が', 'が')],
        [seg('欲しい', 'ほしい', '欲しい')],
        [seg('です', 'です', 'です')],
      ],
      {
        tokenizeWithMecab: async () => [
          mecabToken('私', 0, '名詞'),
          mecabToken('も', 1, '助詞'),
          mecabToken(' ', 2, '記号'),
          mecabToken('あの', 3, '連体詞'),
          mecabToken('仮面', 5, '名詞'),
          mecabToken('が', 7, '助詞'),
          mecabToken('欲しい', 8, '形容詞'),
          mecabToken('です', 11, '助動詞'),
        ],
        isKnownWord: (text) => text === '私' || text === 'あの' || text === '欲しい',
      },
    ),
  );

  const targets = result.tokens?.filter((token) => token.isNPlusOneTarget) ?? [];
  for (const [surface, pos1] of [
    ['が', '助詞'],
    ['です', '助動詞'],
  ]) {
    const token = result.tokens?.find((candidate) => candidate.surface === surface);
    assert.deepEqual(
      {
        pos1: token?.pos1,
        isKnown: token?.isKnown,
        isNPlusOneTarget: token?.isNPlusOneTarget,
        jlptLevel: token?.jlptLevel,
        frequencyRank: token?.frequencyRank,
      },
      {
        pos1,
        isKnown: false,
        isNPlusOneTarget: false,
        jlptLevel: undefined,
        frequencyRank: undefined,
      },
      surface,
    );
  }
  assert.equal(targets.length, 1);
  assert.equal(targets[0]?.surface, '仮面');
});

test('tokenizeSubtitle preserves merged token frequency when MeCab positions cross a newline gap', async () => {
  const parserWindow = createYomitanParserWindow({
    respond: scanTokens([
      { surface: 'X', reading: 'えっくす' },
      { surface: '陰に', reading: 'いんに', startPos: 2 },
      { surface: '潜み', reading: 'ひそ', headword: '潜む' },
    ]),
    frequencies: (pairs) =>
      requestsPair(pairs, '陰に', 'いんに')
        ? [
            yomitanFrequency('陰に', 'いんに', 5702, {
              dictionary: 'JPDBv2㋕',
              displayValueParsed: false,
            }),
          ]
        : [],
  });

  const deps = createTokenizerDepsRuntime(
    makeRuntimeOptions({
      getYomitanExt: () => ({ id: 'dummy-ext' }) as Electron.Extension,
      getYomitanParserWindow: () => parserWindow,
      getFrequencyDictionaryEnabled: () => true,
      getMecabTokenizer: () => ({
        tokenize: async () => [
          mecabWord('X', PartOfSpeech.noun, '名詞', '一般', 'エックス'),
          mecabWord('陰', PartOfSpeech.noun, '名詞', '一般', 'カゲ'),
          mecabWord('に', PartOfSpeech.particle, '助詞', '格助詞', 'ニ', { pos3: '一般' }),
          mecabWord('潜み', PartOfSpeech.verb, '動詞', '自立', 'ヒソミ', {
            headword: '潜む',
            inflectionType: '五段・マ行',
            inflectionForm: '連用形',
          }),
        ],
      }),
    }),
  );

  const result = await tokenizeSubtitle('X\n陰に潜み', deps);

  assert.equal(result.tokens?.[1]?.surface, '陰に');
  assert.equal(result.tokens?.[1]?.pos1, '名詞|助詞');
  assert.equal(result.tokens?.[1]?.pos2, '一般|格助詞');
  assert.equal(result.tokens?.[1]?.frequencyRank, 5702);
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

test('tokenizeSubtitle preserves Yomitan compound token when MeCab components are known', async () => {
  const text = '取り組んでもらいます';
  const result = await tokenizeSubtitle(
    text,
    makeDepsFromYomitanTokens(
      [
        { surface: '取り組んで', reading: 'とりくんで', headword: '取り組む' },
        { surface: 'もらいます', reading: 'もらいます', headword: 'もらう' },
      ],
      {
        isKnownWord: (word) => word === '取る' || word === '組む' || word === 'もらう',
        tokenizeWithMecab: async () => [
          mecabToken('取り組ん', 0, '動詞', '自立', '*'),
          mecabToken('で', 4, '助詞', '接続助詞', '*'),
          mecabToken('もらい', 5, '動詞', '非自立', '*'),
          mecabToken('ます', 8, '助動詞', '*', '*'),
        ],
      },
    ),
  );

  assert.equal(result.text, text);
  assert.equal(result.tokens?.[0]?.surface, '取り組んで');
  assert.equal(result.tokens?.[0]?.headword, '取り組む');
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(result.tokens?.[0]?.pos1, '動詞|助詞');
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

const KAMEN_WORD = mecabWord('仮面', PartOfSpeech.noun, '名詞', '一般', 'カメン');

test('createTokenizerDepsRuntime checks MeCab availability before first tokenizeWithMecab call', async () => {
  let available = false;
  let checkCalls = 0;

  const deps = createTokenizerDepsRuntime(
    makeRuntimeOptions({
      getMecabTokenizer: () => ({
        getStatus: () => ({ available }),
        checkAvailability: async () => {
          checkCalls += 1;
          available = true;
          return true;
        },
        tokenize: async () => (available ? [KAMEN_WORD] : null),
      }),
    }),
  );

  const first = await deps.tokenizeWithMecab('仮面');
  const second = await deps.tokenizeWithMecab('仮面');

  assert.equal(checkCalls, 1);
  assert.equal(first?.[0]?.surface, '仮面');
  assert.equal(second?.[0]?.surface, '仮面');
});

test('createTokenizerDepsRuntime skips known-word lookup for MeCab POS enrichment tokens', async () => {
  let knownWordCalls = 0;

  const deps = createTokenizerDepsRuntime(
    makeRuntimeOptions({
      isKnownWord: () => {
        knownWordCalls += 1;
        return true;
      },
      getMecabTokenizer: () => ({ tokenize: async () => [KAMEN_WORD] }),
    }),
  );

  const tokens = await deps.tokenizeWithMecab('仮面');

  assert.equal(knownWordCalls, 0);
  assert.equal(tokens?.[0]?.isKnown, false);
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
      { tokenizeWithMecab: async () => [mecabToken('は', 0, '助詞', '係助詞', '*')] },
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
        return [mecabToken('猫', 0, '名詞')];
      },
    }),
  );

  assert.equal(result.tokens?.[0]?.frequencyRank, 7);
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(knownCalls, 0);
  assert.equal(mecabCalls, 1);
  assert.equal(frequencyCalls, 1);
});

test('tokenizeSubtitle keeps excluded interjections hoverable while clearing annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    'ぐはっ 猫',
    makeDepsFromScanningParser([[seg('ぐはっ', 'ぐはっ', 'ぐはっ')], [seg('猫', 'ねこ', '猫')]], {
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === '猫' ? 11 : 17),
      getJlptLevel: (text) => (text === '猫' ? 'N5' : null),
      tokenizeWithMecab: async () => [
        mecabToken('ぐはっ', 0, '感動詞'),
        mecabToken('猫', 4, '名詞'),
      ],
    }),
  );

  assert.equal(result.text, 'ぐはっ 猫');
  assert.deepEqual(
    result.tokens?.map(({ surface, headword, frequencyRank, jlptLevel }) => ({
      surface,
      headword,
      frequencyRank,
      jlptLevel,
    })),
    [
      { surface: 'ぐはっ', headword: 'ぐはっ', frequencyRank: undefined, jlptLevel: undefined },
      { surface: '猫', headword: '猫', frequencyRank: 11, jlptLevel: 'N5' },
    ],
  );
});

test('tokenizeSubtitle keeps standalone grammar-only tokens hoverable while clearing annotation metadata', async () => {
  const result = await tokenizeSubtitle(
    '私はこの猫です',
    makeDepsFromScanningParser(
      [
        [seg('私', 'わたし', '私')],
        [seg('は', 'は', 'は')],
        [seg('この', 'この', 'この')],
        [seg('猫', 'ねこ', '猫')],
        [seg('です', 'です', 'です')],
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '私' ? 50 : text === '猫' ? 11 : 500),
        getJlptLevel: (text) => (text === '私' ? 'N5' : text === '猫' ? 'N5' : null),
        tokenizeWithMecab: async () => [
          mecabToken('私', 0, '名詞', '代名詞'),
          mecabToken('は', 1, '助詞', '係助詞'),
          mecabToken('この', 2, '連体詞'),
          mecabToken('猫', 4, '名詞', '一般'),
          mecabToken('です', 5, '助動詞'),
        ],
      },
    ),
  );

  assert.equal(result.text, '私はこの猫です');
  assert.deepEqual(
    result.tokens?.map(({ surface, headword, frequencyRank, jlptLevel }) => ({
      surface,
      headword,
      frequencyRank,
      jlptLevel,
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

test('tokenizeSubtitle keeps frequency for content-led merged token with trailing colloquial suffixes', async () => {
  const result = await tokenizeSubtitle(
    '張り切ってんじゃ',
    makeDepsFromYomitanTokens(
      [{ surface: '張り切ってん', reading: 'はき', headword: '張り切る' }],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '張り切る' ? 5468 : null),
        tokenizeWithMecab: async () => [
          mecabToken('張り切っ', 0, '動詞', '自立'),
          mecabToken('て', 4, '助詞', '接続助詞'),
          mecabToken('んじゃ', 5, '接続詞', '*'),
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

test('tokenizeSubtitle keeps Yomitan frequency for noun-particle-noun compounds', async () => {
  const result = await tokenizeSubtitle(
    '目の前',
    makeDepsFromYomitanTokens(
      [{ surface: '目の前', reading: 'めのまえ', headword: '目の前', frequencyRank: 581 }],
      {
        getFrequencyDictionaryEnabled: () => true,
        tokenizeWithMecab: async () => [
          mecabToken('目', 0, '名詞', '一般'),
          mecabToken('の', 1, '助詞', '連体化'),
          mecabToken('前', 2, '名詞', '副詞可能'),
        ],
      },
    ),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, '目の前');
  assert.equal(result.tokens?.[0]?.pos1, '名詞|助詞');
  assert.equal(result.tokens?.[0]?.frequencyRank, 581);
});

test('tokenizeSubtitle skips frequency requests for ranks supplied by the scanner', async () => {
  const deps = makeDepsFromYomitanTokens(
    [{ surface: '猫', reading: 'ねこ', headword: '猫', frequencyRank: 42 }],
    { getFrequencyDictionaryEnabled: () => true },
  );
  const parserWindow = deps.getYomitanParserWindow();
  assert.ok(parserWindow);
  deps.getYomitanParserWindow = () => parserWindow;
  const scripts: string[] = [];
  const execute = parserWindow.webContents.executeJavaScript.bind(parserWindow.webContents);
  parserWindow.webContents.executeJavaScript = async (script) => {
    scripts.push(script);
    return execute(script);
  };
  const result = await tokenizeSubtitle('猫', deps);
  assert.equal(result.tokens?.[0]?.frequencyRank, 42);
  assert.ok(scripts.length > 0);
  assert.equal(scripts.filter((script) => script.includes('getTermFrequencies')).length, 0);
});

test('tokenizeSubtitle keeps frequency for ordinal prefix-noun tokens', async () => {
  const result = await tokenizeSubtitle(
    '第二走者',
    makeDepsFromYomitanTokens(
      [
        { surface: '第二', reading: 'だいに', headword: '第二' },
        { surface: '走者', reading: 'そうしゃ', headword: '走者' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === '第二' ? 1820 : text === '走者' ? 41555 : null),
        tokenizeWithMecab: async () => [
          mecabToken('第', 0, '接頭詞', '数接続'),
          mecabToken('二', 1, '名詞', '数'),
          mecabToken('走者', 2, '名詞', '一般'),
        ],
        getMinSentenceWordsForNPlusOne: () => 1,
      },
    ),
  );

  assert.equal(result.tokens?.[0]?.surface, '第二');
  assert.equal(result.tokens?.[0]?.pos1, '接頭詞|名詞');
  assert.equal(result.tokens?.[0]?.pos2, '数接続|数');
  assert.equal(result.tokens?.[0]?.frequencyRank, 1820);
});

test('tokenizeSubtitle keeps frequency for honorific prefix-noun tokens', async () => {
  const result = await tokenizeSubtitle(
    'ご機嫌が良くない',
    makeDepsFromYomitanTokens(
      [
        { surface: 'ご機嫌', reading: 'ごきげん', headword: 'ご機嫌' },
        { surface: 'が', reading: 'が', headword: 'が' },
        { surface: '良くない', reading: 'よくない', headword: '良い' },
      ],
      {
        getFrequencyDictionaryEnabled: () => true,
        getFrequencyRank: (text) => (text === 'ご機嫌' ? 5484 : null),
        tokenizeWithMecab: async () => [
          mecabToken('ご', 0, '接頭詞', '名詞接続'),
          mecabToken('機嫌', 1, '名詞', '一般'),
          mecabToken('が', 3, '助詞', '格助詞'),
          mecabToken('良く', 4, '形容詞', '自立'),
          mecabToken('ない', 6, '助動詞', '*'),
        ],
        getMinSentenceWordsForNPlusOne: () => 1,
      },
    ),
  );

  assert.equal(result.tokens?.[0]?.surface, 'ご機嫌');
  assert.equal(result.tokens?.[0]?.pos1, '接頭詞|名詞');
  assert.equal(result.tokens?.[0]?.pos2, '名詞接続|一般');
  assert.equal(result.tokens?.[0]?.frequencyRank, 5484);
});
