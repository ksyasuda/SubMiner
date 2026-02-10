import test from "node:test";
import assert from "node:assert/strict";
import { PartOfSpeech } from "../../types";
import { createTokenizerDepsRuntimeService } from "./tokenizer-deps-runtime-service";

test("createTokenizerDepsRuntimeService tokenizes with mecab and merge", async () => {
  let parserWindow: any = null;
  let readyPromise: Promise<void> | null = null;
  let initPromise: Promise<boolean> | null = null;

  const deps = createTokenizerDepsRuntimeService({
    getYomitanExt: () => null,
    getYomitanParserWindow: () => parserWindow,
    setYomitanParserWindow: (window) => {
      parserWindow = window;
    },
    getYomitanParserReadyPromise: () => readyPromise,
    setYomitanParserReadyPromise: (promise) => {
      readyPromise = promise;
    },
    getYomitanParserInitPromise: () => initPromise,
    setYomitanParserInitPromise: (promise) => {
      initPromise = promise;
    },
    getMecabTokenizer: () => ({
      tokenize: async () => [
        {
          word: "猫",
          partOfSpeech: PartOfSpeech.noun,
          pos1: "名詞",
          pos2: "一般",
          pos3: "",
          pos4: "",
          inflectionType: "",
          inflectionForm: "",
          headword: "猫",
          katakanaReading: "ネコ",
          pronunciation: "ネコ",
        },
      ],
    }),
  });

  const merged = await deps.tokenizeWithMecab("猫");
  assert.ok(Array.isArray(merged));
  assert.equal(merged?.length, 1);
  assert.equal(merged?.[0]?.surface, "猫");
});
