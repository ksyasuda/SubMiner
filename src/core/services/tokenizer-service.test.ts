import test from "node:test";
import assert from "node:assert/strict";
import { PartOfSpeech } from "../../types";
import {
  createTokenizerDepsRuntimeService,
  TokenizerServiceDeps,
  TokenizerDepsRuntimeOptions,
  tokenizeSubtitleService,
} from "./tokenizer-service";

function makeDeps(
  overrides: Partial<TokenizerServiceDeps> = {},
): TokenizerServiceDeps {
  return {
    getYomitanExt: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    isKnownWord: () => false,
    getKnownWordMatchMode: () => "headword",
    getJlptLevel: () => null,
    tokenizeWithMecab: async () => null,
    ...overrides,
  };
}

function makeDepsFromMecabTokenizer(
  tokenize: (text: string) => Promise<import("../../types").Token[] | null>,
  overrides: Partial<TokenizerDepsRuntimeOptions> = {},
): TokenizerServiceDeps {
  return createTokenizerDepsRuntimeService({
    getYomitanExt: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => {},
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => {},
    isKnownWord: () => false,
    getKnownWordMatchMode: () => "headword",
    getMecabTokenizer: () => ({
      tokenize,
    }),
    getJlptLevel: () => null,
    ...overrides,
  });
}

test("tokenizeSubtitleService assigns JLPT level to parsed Yomitan tokens", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "猫",
                    reading: "ねこ",
                    headwords: [[{ term: "猫" }]],
                  },
                  {
                    text: "です",
                    reading: "です",
                    headwords: [[{ term: "です" }]],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      tokenizeWithMecab: async () => null,
      getJlptLevel: (text) => (text === "猫" ? "N5" : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, "N5");
});

test("tokenizeSubtitleService caches JLPT lookups across repeated tokens", async () => {
  let lookupCalls = 0;
  const result = await tokenizeSubtitleService(
    "猫猫",
    makeDepsFromMecabTokenizer(async () => [
      {
        word: "猫",
        partOfSpeech: PartOfSpeech.noun,
        pos1: "",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "猫",
        katakanaReading: "ネコ",
        pronunciation: "ネコ",
      },
      {
        word: "猫",
        partOfSpeech: PartOfSpeech.noun,
        pos1: "",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "猫",
        katakanaReading: "ネコ",
        pronunciation: "ネコ",
      },
    ], {
      getJlptLevel: (text) => {
        lookupCalls += 1;
        return text === "猫" ? "N5" : null;
      },
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(lookupCalls, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, "N5");
  assert.equal(result.tokens?.[1]?.jlptLevel, "N5");
});

test("tokenizeSubtitleService leaves JLPT unset for non-matching tokens", async () => {
  const result = await tokenizeSubtitleService(
    "猫",
    makeDepsFromMecabTokenizer(async () => [
      {
        word: "猫",
        partOfSpeech: PartOfSpeech.noun,
        pos1: "",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "猫",
        katakanaReading: "ネコ",
        pronunciation: "ネコ",
      },
    ], {
      getJlptLevel: () => null,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test("tokenizeSubtitleService skips JLPT lookups when disabled", async () => {
  let lookupCalls = 0;
  const result = await tokenizeSubtitleService(
      "猫です",
      makeDeps({
      tokenizeWithMecab: async () => [
        {
          headword: "猫",
          surface: "猫",
          reading: "ネコ",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getJlptLevel: () => {
        lookupCalls += 1;
        return "N5";
      },
      getJlptEnabled: () => false,
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
  assert.equal(lookupCalls, 0);
});

test("tokenizeSubtitleService applies frequency dictionary ranks", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      tokenizeWithMecab: async () => [
        {
          headword: "猫",
          surface: "猫",
          reading: "ネコ",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: "です",
          surface: "です",
          reading: "デス",
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.bound_auxiliary,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getFrequencyRank: (text) => (text === "猫" ? 23 : 1200),
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.frequencyRank, 23);
  assert.equal(result.tokens?.[1]?.frequencyRank, 1200);
});

test("tokenizeSubtitleService uses all Yomitan headword candidates for frequency lookup", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "猫です",
                    reading: "ねこです",
                    headwords: [
                      [{ term: "猫です" }],
                      [{ term: "猫" }],
                    ],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyRank: (text) => (text === "猫" ? 40 : text === "猫です" ? 1200 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 40);
});

test("tokenizeSubtitleService prefers exact headword frequency over surface/reading when available", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "猫",
                    reading: "ねこ",
                    headwords: [[{ term: "ネコ" }]],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyRank: (text) => (text === "猫" ? 1200 : text === "ネコ" ? 8 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 8);
});

test("tokenizeSubtitleService keeps no frequency when only reading matches and headword candidates miss", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "猫",
                    reading: "ねこ",
                    headwords: [[{ term: "猫です" }]],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyRank: (text) => (text === "ねこ" ? 77 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test("tokenizeSubtitleService ignores invalid frequency ranks and takes best valid headword candidate", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "猫です",
                    reading: "ねこです",
                    headwords: [
                      [{ term: "猫" }],
                      [{ term: "猫です" }],
                    ],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyRank: (text) => (text === "猫" ? Number.NaN : text === "猫です" ? 500 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 500);
});

test("tokenizeSubtitleService handles real-word frequency candidates and prefers most frequent term", async () => {
  const result = await tokenizeSubtitleService(
    "昨日",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "昨日",
                    reading: "きのう",
                    headwords: [
                      [{ term: "昨日" }],
                      [{ term: "きのう" }],
                    ],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyRank: (text) => (text === "きのう" ? 120 : text === "昨日" ? 40 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 40);
});

test("tokenizeSubtitleService ignores candidates with no dictionary rank when higher-frequency candidate exists", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "猫",
                    reading: "ねこ",
                    headwords: [
                      [{ term: "猫" }],
                      [{ term: "猫です" }],
                      [{ term: "unknown-term" }],
                    ],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyRank: (text) => (text === "unknown-term" ? -1 : text === "猫" ? 88 : text === "猫です" ? 9000 : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.frequencyRank, 88);
});

test("tokenizeSubtitleService ignores frequency lookup failures", async () => {
  const result = await tokenizeSubtitleService(
    "猫",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      tokenizeWithMecab: async () => [
        {
          headword: "猫",
          surface: "猫",
          reading: "ネコ",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getFrequencyRank: () => {
        throw new Error("frequency lookup unavailable");
      },
    }),
  );

  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
});

test("tokenizeSubtitleService ignores invalid frequency ranks", async () => {
  const result = await tokenizeSubtitleService(
    "猫",
    makeDeps({
      getFrequencyDictionaryEnabled: () => true,
      tokenizeWithMecab: async () => [
        {
          headword: "猫",
          surface: "猫",
          reading: "ネコ",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          headword: "です",
          surface: "です",
          reading: "デス",
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.bound_auxiliary,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getFrequencyRank: (text) => {
        if (text === "猫") return Number.NaN;
        if (text === "です") return -1;
        return 100;
      },
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
});

test("tokenizeSubtitleService skips frequency lookups when disabled", async () => {
  let frequencyCalls = 0;
  const result = await tokenizeSubtitleService(
    "猫",
    makeDeps({
      getFrequencyDictionaryEnabled: () => false,
      tokenizeWithMecab: async () => [
        {
          headword: "猫",
          surface: "猫",
          reading: "ネコ",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
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

test("tokenizeSubtitleService skips JLPT level for excluded demonstratives", async () => {
  const result = await tokenizeSubtitleService(
    "この",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "この",
                    reading: "この",
                    headwords: [[{ term: "この" }]],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      tokenizeWithMecab: async () => null,
      getJlptLevel: (text) => (text === "この" ? "N5" : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test("tokenizeSubtitleService skips JLPT level for repeated kana SFX", async () => {
  const result = await tokenizeSubtitleService(
    "ああ",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "ああ",
                    reading: "ああ",
                    headwords: [[{ term: "ああ" }]],
                  },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      tokenizeWithMecab: async () => null,
      getJlptLevel: (text) => (text === "ああ" ? "N5" : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test("tokenizeSubtitleService assigns JLPT level to mecab tokens", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDepsFromMecabTokenizer(async () => [
      {
        word: "猫",
        partOfSpeech: PartOfSpeech.noun,
        pos1: "",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "猫",
        katakanaReading: "ネコ",
        pronunciation: "ネコ",
      },
    ], {
      getJlptLevel: (text) => (text === "猫" ? "N4" : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.jlptLevel, "N4");
});

test("tokenizeSubtitleService skips JLPT level for mecab tokens marked as ineligible", async () => {
  const result = await tokenizeSubtitleService(
    "は",
    makeDepsFromMecabTokenizer(async () => [
      {
        word: "は",
        partOfSpeech: PartOfSpeech.particle,
        pos1: "助詞",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "は",
        katakanaReading: "ハ",
        pronunciation: "ハ",
      },
    ], {
      getJlptLevel: (text) => (text === "は" ? "N5" : null),
    }),
  );

  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.pos1, "助詞");
  assert.equal(result.tokens?.[0]?.jlptLevel, undefined);
});

test("tokenizeSubtitleService returns null tokens for empty normalized text", async () => {
  const result = await tokenizeSubtitleService(" \\n  ", makeDeps());
  assert.deepEqual(result, { text: " \\n  ", tokens: null });
});

test("tokenizeSubtitleService normalizes newlines before mecab fallback", async () => {
  let tokenizeInput = "";
  const result = await tokenizeSubtitleService(
    "猫\\Nです\nね",
    makeDeps({
      tokenizeWithMecab: async (text) => {
        tokenizeInput = text;
        return [
        {
            surface: "猫ですね",
            reading: "ネコデスネ",
            headword: "猫ですね",
            startPos: 0,
            endPos: 4,
            partOfSpeech: PartOfSpeech.other,
            isMerged: true,
            isKnown: false,
            isNPlusOneTarget: false,
          },
        ];
      },
    }),
  );

  assert.equal(tokenizeInput, "猫 です ね");
  assert.equal(result.text, "猫\nです\nね");
  assert.equal(result.tokens?.[0]?.surface, "猫ですね");
});

test("tokenizeSubtitleService falls back to mecab tokens when available", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      tokenizeWithMecab: async () => [
        {
          surface: "猫",
          reading: "ネコ",
          headword: "猫",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
    }),
  );

  assert.equal(result.text, "猫です");
  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, "猫");
});

test("tokenizeSubtitleService returns null tokens when mecab throws", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      tokenizeWithMecab: async () => {
        throw new Error("mecab failed");
      },
    }),
  );

  assert.deepEqual(result, { text: "猫です", tokens: null });
});

test("tokenizeSubtitleService uses Yomitan parser result when available", async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async () => [
        {
          source: "scanning-parser",
          index: 0,
          content: [
            [
              {
                text: "猫",
                reading: "ねこ",
                headwords: [[{ term: "猫" }]],
              },
              {
                text: "です",
                reading: "です",
              },
            ],
          ],
        },
      ],
    },
  } as unknown as Electron.BrowserWindow;

  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => parserWindow,
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.text, "猫です");
  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, "猫");
  assert.equal(result.tokens?.[0]?.reading, "ねこ");
  assert.equal(result.tokens?.[0]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.surface, "です");
  assert.equal(result.tokens?.[1]?.reading, "です");
  assert.equal(result.tokens?.[1]?.isKnown, false);
});

test("tokenizeSubtitleService prefers mecab parser tokens when scanning parser returns one token", async () => {
  const result = await tokenizeSubtitleService(
    "俺は小園にいきたい",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  {
                    text: "俺は小園にいきたい",
                    reading: "おれは小園にいきたい",
                    headwords: [[{ term: "俺は小園にいきたい" }]],
                  },
                ],
              ],
            },
            {
              source: "mecab",
              index: 0,
              content: [
                [{ text: "俺", reading: "おれ", headwords: [[{ term: "俺" }]] }],
                [{ text: "は", reading: "は", headwords: [[{ term: "は" }]] }],
                [{ text: "小園", reading: "おうえん", headwords: [[{ term: "小園" }]] }],
                [{ text: "に", reading: "に", headwords: [[{ term: "に" }]] }],
                [{ text: "いきたい", reading: "いきたい", headwords: [[{ term: "いきたい" }]] }],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyDictionaryEnabled: () => true,
      tokenizeWithMecab: async () => null,
      getFrequencyRank: (text) =>
        text === "小園" ? 25 : text === "いきたい" ? 1500 : null,
    }),
  );

  assert.equal(result.tokens?.length, 5);
  assert.equal(result.tokens?.map((token) => token.surface).join(","), "俺,は,小園,に,いきたい");
  assert.equal(result.tokens?.[2]?.surface, "小園");
  assert.equal(result.tokens?.[2]?.frequencyRank, 25);
});

test("tokenizeSubtitleService keeps scanning parser tokens when they are already split", async () => {
  const result = await tokenizeSubtitleService(
    "小園に行きたい",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [{ text: "小園", reading: "おうえん", headwords: [[{ term: "小園" }]] }],
                [{ text: "に", reading: "に", headwords: [[{ term: "に" }]] }],
                [{ text: "行きたい", reading: "いきたい", headwords: [[{ term: "行きたい" }]] }],
              ],
            },
            {
              source: "mecab",
              index: 0,
              content: [
                [{ text: "小", reading: "お", headwords: [[{ term: "小" }]] }],
                [{ text: "園", reading: "えん", headwords: [[{ term: "園" }]] }],
                [{ text: "に", reading: "に", headwords: [[{ term: "に" }]] }],
                [{ text: "行き", reading: "いき", headwords: [[{ term: "行き" }]] }],
                [{ text: "たい", reading: "たい", headwords: [[{ term: "たい" }]] }],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === "小園" ? 20 : null),
      tokenizeWithMecab: async () => null,
    }),
  );

  assert.equal(result.tokens?.length, 3);
  assert.equal(
    result.tokens?.map((token) => token.surface).join(","),
    "小園,に,行きたい",
  );
  assert.equal(result.tokens?.[0]?.frequencyRank, 20);
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
  assert.equal(result.tokens?.[2]?.frequencyRank, undefined);
});

test("tokenizeSubtitleService still assigns frequency to non-known Yomitan tokens", async () => {
  const result = await tokenizeSubtitleService(
    "小園に",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => ({
        isDestroyed: () => false,
        webContents: {
          executeJavaScript: async () => [
            {
              source: "scanning-parser",
              index: 0,
              content: [
                [
                  { text: "小園", reading: "おうえん", headwords: [[{ term: "小園" }]] },
                ],
                [
                  { text: "に", reading: "に", headwords: [[{ term: "に" }]] },
                ],
              ],
            },
          ],
        },
      } as unknown as Electron.BrowserWindow),
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === "小園" ? 75 : text === "に" ? 3000 : null),
      isKnownWord: (text) => text === "小園",
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.isKnown, true);
  assert.equal(result.tokens?.[0]?.frequencyRank, 75);
  assert.equal(result.tokens?.[1]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.frequencyRank, undefined);
});

test("tokenizeSubtitleService marks tokens as known using callback", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDepsFromMecabTokenizer(async () => [
      {
        word: "猫",
        partOfSpeech: PartOfSpeech.noun,
        pos1: "",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "猫",
        katakanaReading: "ネコ",
        pronunciation: "ネコ",
      },
    ], {
      isKnownWord: (text) => text === "猫",
    }),
  );

  assert.equal(result.text, "猫です");
  assert.equal(result.tokens?.[0]?.isKnown, true);
});

test("tokenizeSubtitleService still assigns frequency rank to non-known tokens", async () => {
  const result = await tokenizeSubtitleService(
    "既知未知",
    makeDeps({
      tokenizeWithMecab: async () => [
        {
          surface: "既知",
          reading: "キチ",
          partOfSpeech: PartOfSpeech.noun,
          pos1: "",
          pos2: "",
          pos3: "",
          pos4: "",
          inflectionType: "",
          inflectionForm: "",
          headword: "既知",
          katakanaReading: "キチ",
          pronunciation: "キチ",
          startPos: 0,
          endPos: 2,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: "未知",
          reading: "ミチ",
          partOfSpeech: PartOfSpeech.noun,
          pos1: "",
          pos2: "",
          pos3: "",
          pos4: "",
          inflectionType: "",
          inflectionForm: "",
          headword: "未知",
          katakanaReading: "ミチ",
          pronunciation: "ミチ",
          startPos: 2,
          endPos: 4,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank: (text) => (text === "既知" ? 20 : text === "未知" ? 30 : null),
      isKnownWord: (text) => text === "既知",
    }),
  );

  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.isKnown, true);
  assert.equal(result.tokens?.[0]?.frequencyRank, 20);
  assert.equal(result.tokens?.[1]?.isKnown, false);
  assert.equal(result.tokens?.[1]?.frequencyRank, 30);
});

test("tokenizeSubtitleService selects one N+1 target token", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      tokenizeWithMecab: async () => [
        {
          surface: "私",
          reading: "ワタシ",
          headword: "私",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: "犬",
          reading: "イヌ",
          headword: "犬",
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
      isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
      getMinSentenceWordsForNPlusOne: () => 2,
    }),
  );

  const targets = result.tokens?.filter((token) => token.isNPlusOneTarget) ?? [];
  assert.equal(targets.length, 1);
  assert.equal(targets[0]?.surface, "犬");
});

test("tokenizeSubtitleService does not mark target when sentence has multiple candidates", async () => {
  const result = await tokenizeSubtitleService(
    "猫犬",
    makeDeps({
      tokenizeWithMecab: async () => [
        {
          surface: "猫",
          reading: "ネコ",
          headword: "猫",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
        {
          surface: "犬",
          reading: "イヌ",
          headword: "犬",
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
    }),
  );

  assert.equal(
    result.tokens?.some((token) => token.isNPlusOneTarget),
    false,
  );
});

test("tokenizeSubtitleService applies N+1 target marking to Yomitan results", async () => {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async () => [
        {
          source: "scanning-parser",
          index: 0,
      content: [
        [
          {
            text: "猫",
            reading: "ねこ",
            headwords: [[{ term: "猫" }]],
          },
        ],
        [
          {
            text: "です",
            reading: "です",
            headwords: [[{ term: "です" }]],
          },
        ],
      ],
    },
      ],
    },
  } as unknown as Electron.BrowserWindow;

  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      getYomitanExt: () => ({ id: "dummy-ext" } as any),
      getYomitanParserWindow: () => parserWindow,
      tokenizeWithMecab: async () => null,
      isKnownWord: (text) => text === "です",
      getMinSentenceWordsForNPlusOne: () => 2,
    }),
  );

  assert.equal(result.text, "猫です");
  assert.equal(result.tokens?.length, 2);
  assert.equal(result.tokens?.[0]?.surface, "猫");
  assert.equal(result.tokens?.[0]?.isNPlusOneTarget, true);
  assert.equal(result.tokens?.[1]?.isNPlusOneTarget, false);
});

test("tokenizeSubtitleService does not color 1-2 word sentences by default", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDeps({
      tokenizeWithMecab: async () => [
        {
          surface: "私",
          reading: "ワタシ",
          headword: "私",
          startPos: 0,
          endPos: 1,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: true,
          isNPlusOneTarget: false,
        },
        {
          surface: "犬",
          reading: "イヌ",
          headword: "犬",
          startPos: 1,
          endPos: 2,
          partOfSpeech: PartOfSpeech.noun,
          isMerged: false,
          isKnown: false,
          isNPlusOneTarget: false,
        },
      ],
    }),
  );

  assert.equal(
    result.tokens?.some((token) => token.isNPlusOneTarget),
    false,
  );
});

test("tokenizeSubtitleService checks known words by headword, not surface", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDepsFromMecabTokenizer(async () => [
      {
        word: "猫",
        partOfSpeech: PartOfSpeech.noun,
        pos1: "",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "猫です",
        katakanaReading: "ネコ",
        pronunciation: "ネコ",
      },
    ], {
      isKnownWord: (text) => text === "猫です",
    }),
  );

  assert.equal(result.text, "猫です");
  assert.equal(result.tokens?.[0]?.isKnown, true);
});

test("tokenizeSubtitleService checks known words by surface when configured", async () => {
  const result = await tokenizeSubtitleService(
    "猫です",
    makeDepsFromMecabTokenizer(async () => [
      {
        word: "猫",
        partOfSpeech: PartOfSpeech.noun,
        pos1: "",
        pos2: "",
        pos3: "",
        pos4: "",
        inflectionType: "",
        inflectionForm: "",
        headword: "猫です",
        katakanaReading: "ネコ",
        pronunciation: "ネコ",
      },
    ], {
      getKnownWordMatchMode: () => "surface",
      isKnownWord: (text) => text === "猫",
    }),
  );

  assert.equal(result.text, "猫です");
  assert.equal(result.tokens?.[0]?.isKnown, true);
});
