import test from "node:test";
import assert from "node:assert/strict";
import { PartOfSpeech } from "../../types";
import { tokenizeSubtitleService, TokenizerServiceDeps } from "./tokenizer-service";

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
    tokenizeWithMecab: async () => null,
    ...overrides,
  };
}

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
});
