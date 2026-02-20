export function createBuildTokenizerDepsMainHandler(deps: {
  getYomitanExt: () => unknown;
  getYomitanParserWindow: () => unknown;
  setYomitanParserWindow: (window: unknown) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  isKnownWord: (text: string) => boolean;
  recordLookup: (hit: boolean) => void;
  getKnownWordMatchMode: () => unknown;
  getMinSentenceWordsForNPlusOne: () => number;
  getJlptLevel: (text: string) => unknown;
  getJlptEnabled: () => boolean;
  getFrequencyDictionaryEnabled: () => boolean;
  getFrequencyRank: (text: string) => unknown;
  getYomitanGroupDebugEnabled: () => boolean;
  getMecabTokenizer: () => unknown;
}) {
  return () => ({
    getYomitanExt: () => deps.getYomitanExt() as never,
    getYomitanParserWindow: () => deps.getYomitanParserWindow() as never,
    setYomitanParserWindow: (window: unknown) => deps.setYomitanParserWindow(window),
    getYomitanParserReadyPromise: () => deps.getYomitanParserReadyPromise() as never,
    setYomitanParserReadyPromise: (promise: Promise<void> | null) =>
      deps.setYomitanParserReadyPromise(promise),
    getYomitanParserInitPromise: () => deps.getYomitanParserInitPromise() as never,
    setYomitanParserInitPromise: (promise: Promise<boolean> | null) =>
      deps.setYomitanParserInitPromise(promise),
    isKnownWord: (text: string) => {
      const hit = deps.isKnownWord(text);
      deps.recordLookup(hit);
      return hit;
    },
    getKnownWordMatchMode: () => deps.getKnownWordMatchMode() as never,
    getMinSentenceWordsForNPlusOne: () => deps.getMinSentenceWordsForNPlusOne(),
    getJlptLevel: (text: string) => deps.getJlptLevel(text) as never,
    getJlptEnabled: () => deps.getJlptEnabled(),
    getFrequencyDictionaryEnabled: () => deps.getFrequencyDictionaryEnabled(),
    getFrequencyRank: (text: string) => deps.getFrequencyRank(text) as never,
    getYomitanGroupDebugEnabled: () => deps.getYomitanGroupDebugEnabled(),
    getMecabTokenizer: () => deps.getMecabTokenizer() as never,
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
