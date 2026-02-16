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
  assert.equal(result.tokens?.length, 1);
  assert.equal(result.tokens?.[0]?.surface, "猫です");
  assert.equal(result.tokens?.[0]?.reading, "ねこです");
  assert.equal(result.tokens?.[0]?.isKnown, false);
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
