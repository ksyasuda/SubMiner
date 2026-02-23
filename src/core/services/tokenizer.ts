import type { BrowserWindow, Extension } from 'electron';
import { mergeTokens } from '../../token-merger';
import { createLogger } from '../../logger';
import {
  MergedToken,
  NPlusOneMatchMode,
  SubtitleData,
  Token,
  FrequencyDictionaryLookup,
  JlptLevel,
} from '../../types';
import { annotateTokens } from './tokenizer/annotation-stage';
import { enrichTokensWithMecabPos1 } from './tokenizer/parser-enrichment-stage';
import { selectYomitanParseTokens } from './tokenizer/parser-selection-stage';
import { requestYomitanParseResults } from './tokenizer/yomitan-parser-runtime';

const logger = createLogger('main:tokenizer');

export interface TokenizerServiceDeps {
  getYomitanExt: () => Extension | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  isKnownWord: (text: string) => boolean;
  getKnownWordMatchMode: () => NPlusOneMatchMode;
  getJlptLevel: (text: string) => JlptLevel | null;
  getJlptEnabled?: () => boolean;
  getFrequencyDictionaryEnabled?: () => boolean;
  getFrequencyRank?: FrequencyDictionaryLookup;
  getMinSentenceWordsForNPlusOne?: () => number;
  getYomitanGroupDebugEnabled?: () => boolean;
  tokenizeWithMecab: (text: string) => Promise<MergedToken[] | null>;
}

interface MecabTokenizerLike {
  tokenize: (text: string) => Promise<Token[] | null>;
  checkAvailability?: () => Promise<boolean>;
  getStatus?: () => { available: boolean };
}

export interface TokenizerDepsRuntimeOptions {
  getYomitanExt: () => Extension | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  isKnownWord: (text: string) => boolean;
  getKnownWordMatchMode: () => NPlusOneMatchMode;
  getJlptLevel: (text: string) => JlptLevel | null;
  getJlptEnabled?: () => boolean;
  getFrequencyDictionaryEnabled?: () => boolean;
  getFrequencyRank?: FrequencyDictionaryLookup;
  getMinSentenceWordsForNPlusOne?: () => number;
  getYomitanGroupDebugEnabled?: () => boolean;
  getMecabTokenizer: () => MecabTokenizerLike | null;
}

export function createTokenizerDepsRuntime(
  options: TokenizerDepsRuntimeOptions,
): TokenizerServiceDeps {
  const checkedMecabTokenizers = new WeakSet<object>();

  return {
    getYomitanExt: options.getYomitanExt,
    getYomitanParserWindow: options.getYomitanParserWindow,
    setYomitanParserWindow: options.setYomitanParserWindow,
    getYomitanParserReadyPromise: options.getYomitanParserReadyPromise,
    setYomitanParserReadyPromise: options.setYomitanParserReadyPromise,
    getYomitanParserInitPromise: options.getYomitanParserInitPromise,
    setYomitanParserInitPromise: options.setYomitanParserInitPromise,
    isKnownWord: options.isKnownWord,
    getKnownWordMatchMode: options.getKnownWordMatchMode,
    getJlptLevel: options.getJlptLevel,
    getJlptEnabled: options.getJlptEnabled,
    getFrequencyDictionaryEnabled: options.getFrequencyDictionaryEnabled,
    getFrequencyRank: options.getFrequencyRank,
    getMinSentenceWordsForNPlusOne: options.getMinSentenceWordsForNPlusOne ?? (() => 3),
    getYomitanGroupDebugEnabled: options.getYomitanGroupDebugEnabled ?? (() => false),
    tokenizeWithMecab: async (text) => {
      const mecabTokenizer = options.getMecabTokenizer();
      if (!mecabTokenizer) {
        return null;
      }

      if (
        typeof mecabTokenizer.checkAvailability === 'function' &&
        typeof mecabTokenizer.getStatus === 'function' &&
        !checkedMecabTokenizers.has(mecabTokenizer as object)
      ) {
        const status = mecabTokenizer.getStatus();
        if (!status.available) {
          await mecabTokenizer.checkAvailability();
        }
        checkedMecabTokenizers.add(mecabTokenizer as object);
      }

      const rawTokens = await mecabTokenizer.tokenize(text);
      if (!rawTokens || rawTokens.length === 0) {
        return null;
      }

      return mergeTokens(rawTokens, options.isKnownWord, options.getKnownWordMatchMode());
    },
  };
}

function logSelectedYomitanGroups(text: string, tokens: MergedToken[]): void {
  if (tokens.length === 0) {
    return;
  }

  logger.info('Selected Yomitan token groups', {
    text,
    tokenCount: tokens.length,
    groups: tokens.map((token, index) => ({
      index,
      surface: token.surface,
      headword: token.headword,
      reading: token.reading,
      startPos: token.startPos,
      endPos: token.endPos,
    })),
  });
}

function getAnnotationOptions(deps: TokenizerServiceDeps): {
  jlptEnabled: boolean;
  frequencyEnabled: boolean;
  minSentenceWordsForNPlusOne: number | undefined;
} {
  return {
    jlptEnabled: deps.getJlptEnabled?.() !== false,
    frequencyEnabled: deps.getFrequencyDictionaryEnabled?.() !== false,
    minSentenceWordsForNPlusOne: deps.getMinSentenceWordsForNPlusOne?.(),
  };
}

function applyAnnotationStage(tokens: MergedToken[], deps: TokenizerServiceDeps): MergedToken[] {
  const options = getAnnotationOptions(deps);

  return annotateTokens(
    tokens,
    {
      isKnownWord: deps.isKnownWord,
      knownWordMatchMode: deps.getKnownWordMatchMode(),
      getJlptLevel: deps.getJlptLevel,
      getFrequencyRank: deps.getFrequencyRank,
    },
    options,
  );
}

async function parseWithYomitanInternalParser(
  text: string,
  deps: TokenizerServiceDeps,
): Promise<MergedToken[] | null> {
  const parseResults = await requestYomitanParseResults(text, deps, logger);
  if (!parseResults) {
    return null;
  }

  const selectedTokens = selectYomitanParseTokens(
    parseResults,
    deps.isKnownWord,
    deps.getKnownWordMatchMode(),
  );
  if (!selectedTokens || selectedTokens.length === 0) {
    return null;
  }

  if (deps.getYomitanGroupDebugEnabled?.() === true) {
    logSelectedYomitanGroups(text, selectedTokens);
  }

  try {
    const mecabTokens = await deps.tokenizeWithMecab(text);
    return enrichTokensWithMecabPos1(selectedTokens, mecabTokens);
  } catch (err) {
    const error = err as Error;
    logger.warn(
      'Failed to enrich Yomitan tokens with MeCab POS:',
      error.message,
      `tokenCount=${selectedTokens.length}`,
      `textLength=${text.length}`,
    );
    return selectedTokens;
  }
}

export async function tokenizeSubtitle(
  text: string,
  deps: TokenizerServiceDeps,
): Promise<SubtitleData> {
  const displayText = text
    .replace(/\r\n/g, '\n')
    .replace(/\\N/g, '\n')
    .replace(/\\n/g, '\n')
    .trim();

  if (!displayText) {
    return { text, tokens: null };
  }

  const tokenizeText = displayText.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();

  const yomitanTokens = await parseWithYomitanInternalParser(tokenizeText, deps);
  if (yomitanTokens && yomitanTokens.length > 0) {
    return {
      text: displayText,
      tokens: applyAnnotationStage(yomitanTokens, deps),
    };
  }

  return { text: displayText, tokens: null };
}
