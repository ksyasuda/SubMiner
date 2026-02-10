import { BrowserWindow, Extension } from "electron";
import { mergeTokens } from "../../token-merger";
import { TokenizerServiceDeps } from "./tokenizer-service";

interface RawTokenLike {}

interface MecabTokenizerLike {
  tokenize: (text: string) => Promise<RawTokenLike[] | null>;
}

export interface TokenizerDepsRuntimeOptions {
  getYomitanExt: () => Extension | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  getMecabTokenizer: () => MecabTokenizerLike | null;
}

export function createTokenizerDepsRuntimeService(
  options: TokenizerDepsRuntimeOptions,
): TokenizerServiceDeps {
  return {
    getYomitanExt: options.getYomitanExt,
    getYomitanParserWindow: options.getYomitanParserWindow,
    setYomitanParserWindow: options.setYomitanParserWindow,
    getYomitanParserReadyPromise: options.getYomitanParserReadyPromise,
    setYomitanParserReadyPromise: options.setYomitanParserReadyPromise,
    getYomitanParserInitPromise: options.getYomitanParserInitPromise,
    setYomitanParserInitPromise: options.setYomitanParserInitPromise,
    tokenizeWithMecab: async (text) => {
      const mecabTokenizer = options.getMecabTokenizer();
      if (!mecabTokenizer) {
        return null;
      }
      const rawTokens = await mecabTokenizer.tokenize(text);
      if (!rawTokens || rawTokens.length === 0) {
        return null;
      }
      return mergeTokens(rawTokens as never);
    },
  };
}
