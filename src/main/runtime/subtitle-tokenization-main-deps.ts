import type { TokenizerDepsRuntimeOptions } from '../../core/services/tokenizer';

type TokenizerMainDeps = TokenizerDepsRuntimeOptions & {
  getJlptEnabled: NonNullable<TokenizerDepsRuntimeOptions['getJlptEnabled']>;
  getFrequencyDictionaryEnabled: NonNullable<
    TokenizerDepsRuntimeOptions['getFrequencyDictionaryEnabled']
  >;
  getFrequencyRank: NonNullable<TokenizerDepsRuntimeOptions['getFrequencyRank']>;
  getMinSentenceWordsForNPlusOne: NonNullable<
    TokenizerDepsRuntimeOptions['getMinSentenceWordsForNPlusOne']
  >;
  getYomitanGroupDebugEnabled: NonNullable<
    TokenizerDepsRuntimeOptions['getYomitanGroupDebugEnabled']
  >;
  recordLookup: (hit: boolean) => void;
};

export function createBuildTokenizerDepsMainHandler(deps: TokenizerMainDeps) {
  return (): TokenizerDepsRuntimeOptions => ({
    getYomitanExt: () => deps.getYomitanExt(),
    getYomitanParserWindow: () => deps.getYomitanParserWindow(),
    setYomitanParserWindow: (window) => deps.setYomitanParserWindow(window),
    getYomitanParserReadyPromise: () => deps.getYomitanParserReadyPromise(),
    setYomitanParserReadyPromise: (promise: Promise<void> | null) =>
      deps.setYomitanParserReadyPromise(promise),
    getYomitanParserInitPromise: () => deps.getYomitanParserInitPromise(),
    setYomitanParserInitPromise: (promise: Promise<boolean> | null) =>
      deps.setYomitanParserInitPromise(promise),
    isKnownWord: (text: string) => {
      const hit = deps.isKnownWord(text);
      deps.recordLookup(hit);
      return hit;
    },
    getKnownWordMatchMode: () => deps.getKnownWordMatchMode(),
    getMinSentenceWordsForNPlusOne: () => deps.getMinSentenceWordsForNPlusOne(),
    getJlptLevel: (text: string) => deps.getJlptLevel(text),
    getJlptEnabled: () => deps.getJlptEnabled(),
    getFrequencyDictionaryEnabled: () => deps.getFrequencyDictionaryEnabled(),
    getFrequencyRank: (text: string) => deps.getFrequencyRank(text),
    getYomitanGroupDebugEnabled: () => deps.getYomitanGroupDebugEnabled(),
    getMecabTokenizer: () => deps.getMecabTokenizer(),
  });
}

export function createCreateMecabTokenizerAndCheckMainHandler<TMecab>(deps: {
  getMecabTokenizer: () => TMecab | null;
  setMecabTokenizer: (tokenizer: TMecab) => void;
  createMecabTokenizer: () => TMecab;
  checkAvailability: (tokenizer: TMecab) => Promise<unknown>;
}) {
  return async (): Promise<void> => {
    let tokenizer = deps.getMecabTokenizer();
    if (!tokenizer) {
      tokenizer = deps.createMecabTokenizer();
      deps.setMecabTokenizer(tokenizer);
    }
    await deps.checkAvailability(tokenizer);
  };
}

export function createPrewarmSubtitleDictionariesMainHandler(deps: {
  ensureJlptDictionaryLookup: () => Promise<void>;
  ensureFrequencyDictionaryLookup: () => Promise<void>;
}) {
  return async (): Promise<void> => {
    await Promise.all([deps.ensureJlptDictionaryLookup(), deps.ensureFrequencyDictionaryLookup()]);
  };
}
