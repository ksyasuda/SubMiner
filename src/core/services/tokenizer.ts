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
import { selectYomitanParseTokens } from './tokenizer/parser-selection-stage';
import {
  requestYomitanParseResults,
  requestYomitanTermFrequencies,
} from './tokenizer/yomitan-parser-runtime';

const logger = createLogger('main:tokenizer');

type MecabTokenEnrichmentFn = (
  tokens: MergedToken[],
  mecabTokens: MergedToken[] | null,
) => Promise<MergedToken[]>;

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
  getNPlusOneEnabled?: () => boolean;
  getJlptEnabled?: () => boolean;
  getFrequencyDictionaryEnabled?: () => boolean;
  getFrequencyRank?: FrequencyDictionaryLookup;
  getMinSentenceWordsForNPlusOne?: () => number;
  getYomitanGroupDebugEnabled?: () => boolean;
  tokenizeWithMecab: (text: string) => Promise<MergedToken[] | null>;
  enrichTokensWithMecab?: MecabTokenEnrichmentFn;
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
  getNPlusOneEnabled?: () => boolean;
  getJlptEnabled?: () => boolean;
  getFrequencyDictionaryEnabled?: () => boolean;
  getFrequencyRank?: FrequencyDictionaryLookup;
  getMinSentenceWordsForNPlusOne?: () => number;
  getYomitanGroupDebugEnabled?: () => boolean;
  getMecabTokenizer: () => MecabTokenizerLike | null;
}

interface TokenizerAnnotationOptions {
  nPlusOneEnabled: boolean;
  jlptEnabled: boolean;
  frequencyEnabled: boolean;
  minSentenceWordsForNPlusOne: number | undefined;
}

let parserEnrichmentWorkerRuntimeModulePromise:
  | Promise<typeof import('./tokenizer/parser-enrichment-worker-runtime')>
  | null = null;
let annotationStageModulePromise: Promise<typeof import('./tokenizer/annotation-stage')> | null = null;
let parserEnrichmentFallbackModulePromise:
  | Promise<typeof import('./tokenizer/parser-enrichment-stage')>
  | null = null;

function getKnownWordLookup(deps: TokenizerServiceDeps, options: TokenizerAnnotationOptions): (text: string) => boolean {
  if (!options.nPlusOneEnabled) {
    return () => false;
  }
  return deps.isKnownWord;
}

function needsMecabPosEnrichment(options: TokenizerAnnotationOptions): boolean {
  return options.jlptEnabled || options.frequencyEnabled;
}

function hasAnyAnnotationEnabled(options: TokenizerAnnotationOptions): boolean {
  return options.nPlusOneEnabled || options.jlptEnabled || options.frequencyEnabled;
}

async function enrichTokensWithMecabAsync(
  tokens: MergedToken[],
  mecabTokens: MergedToken[] | null,
): Promise<MergedToken[]> {
  if (!parserEnrichmentWorkerRuntimeModulePromise) {
    parserEnrichmentWorkerRuntimeModulePromise = import('./tokenizer/parser-enrichment-worker-runtime');
  }

  try {
    const runtime = await parserEnrichmentWorkerRuntimeModulePromise;
    return await runtime.enrichTokensWithMecabPos1Async(tokens, mecabTokens);
  } catch {
    if (!parserEnrichmentFallbackModulePromise) {
      parserEnrichmentFallbackModulePromise = import('./tokenizer/parser-enrichment-stage');
    }
    const fallback = await parserEnrichmentFallbackModulePromise;
    return fallback.enrichTokensWithMecabPos1(tokens, mecabTokens);
  }
}

async function applyAnnotationStage(
  tokens: MergedToken[],
  deps: TokenizerServiceDeps,
  options: TokenizerAnnotationOptions,
): Promise<MergedToken[]> {
  if (!hasAnyAnnotationEnabled(options)) {
    return tokens;
  }

  if (!annotationStageModulePromise) {
    annotationStageModulePromise = import('./tokenizer/annotation-stage');
  }

  const annotationStage = await annotationStageModulePromise;
  return annotationStage.annotateTokens(
    tokens,
    {
      isKnownWord: getKnownWordLookup(deps, options),
      knownWordMatchMode: deps.getKnownWordMatchMode(),
      getJlptLevel: deps.getJlptLevel,
      getFrequencyRank: deps.getFrequencyRank,
    },
    options,
  );
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
    getNPlusOneEnabled: options.getNPlusOneEnabled,
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

      const isKnownWordLookup = options.getNPlusOneEnabled?.() === false ? () => false : options.isKnownWord;
      return mergeTokens(rawTokens, isKnownWordLookup, options.getKnownWordMatchMode());
    },
    enrichTokensWithMecab: async (tokens, mecabTokens) =>
      enrichTokensWithMecabAsync(tokens, mecabTokens),
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

function normalizePositiveFrequencyRank(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    return null;
  }
  return Math.max(1, Math.floor(value));
}

function resolveFrequencyLookupText(token: MergedToken): string {
  if (token.headword && token.headword.length > 0) {
    return token.headword;
  }
  if (token.reading && token.reading.length > 0) {
    return token.reading;
  }
  return token.surface;
}

function applyYomitanFrequencyRanks(
  tokens: MergedToken[],
  frequencies: ReadonlyArray<{ term: string; frequency: number }>,
): MergedToken[] {
  if (tokens.length === 0 || frequencies.length === 0) {
    return tokens;
  }

  const rankByTerm = new Map<string, number>();
  for (const frequency of frequencies) {
    const normalizedTerm = frequency.term.trim();
    const rank = normalizePositiveFrequencyRank(frequency.frequency);
    if (!normalizedTerm || rank === null) {
      continue;
    }
    const current = rankByTerm.get(normalizedTerm);
    if (current === undefined || rank < current) {
      rankByTerm.set(normalizedTerm, rank);
    }
  }

  if (rankByTerm.size === 0) {
    return tokens;
  }

  return tokens.map((token) => {
    const lookupText = resolveFrequencyLookupText(token).trim();
    if (!lookupText) {
      return token;
    }
    const rank = rankByTerm.get(lookupText);
    if (rank === undefined) {
      return token;
    }
    return {
      ...token,
      frequencyRank: rank,
    };
  });
}

function getAnnotationOptions(deps: TokenizerServiceDeps): TokenizerAnnotationOptions {
  return {
    nPlusOneEnabled: deps.getNPlusOneEnabled?.() !== false,
    jlptEnabled: deps.getJlptEnabled?.() !== false,
    frequencyEnabled: deps.getFrequencyDictionaryEnabled?.() !== false,
    minSentenceWordsForNPlusOne: deps.getMinSentenceWordsForNPlusOne?.(),
  };
}

async function parseWithYomitanInternalParser(
  text: string,
  deps: TokenizerServiceDeps,
  options: TokenizerAnnotationOptions,
): Promise<MergedToken[] | null> {
  const parseResults = await requestYomitanParseResults(text, deps, logger);
  if (!parseResults) {
    return null;
  }

  const selectedTokens = selectYomitanParseTokens(
    parseResults,
    getKnownWordLookup(deps, options),
    deps.getKnownWordMatchMode(),
  );
  if (!selectedTokens || selectedTokens.length === 0) {
    return null;
  }

  if (deps.getYomitanGroupDebugEnabled?.() === true) {
    logSelectedYomitanGroups(text, selectedTokens);
  }

  let tokensWithFrequency = selectedTokens;
  if (options.frequencyEnabled) {
    const termReadingList = selectedTokens.map((token) => ({
      term: resolveFrequencyLookupText(token),
      reading: token.reading && token.reading.trim().length > 0 ? token.reading.trim() : null,
    }));
    const yomitanFrequencies = await requestYomitanTermFrequencies(termReadingList, deps, logger);
    tokensWithFrequency = applyYomitanFrequencyRanks(selectedTokens, yomitanFrequencies);
  }

  if (!needsMecabPosEnrichment(options)) {
    return tokensWithFrequency;
  }

  try {
    const mecabTokens = await deps.tokenizeWithMecab(text);
    const enrichTokensWithMecab = deps.enrichTokensWithMecab ?? enrichTokensWithMecabAsync;
    return await enrichTokensWithMecab(tokensWithFrequency, mecabTokens);
  } catch (err) {
    const error = err as Error;
    logger.warn(
      'Failed to enrich Yomitan tokens with MeCab POS:',
      error.message,
      `tokenCount=${selectedTokens.length}`,
      `textLength=${text.length}`,
    );
    return tokensWithFrequency;
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
  const annotationOptions = getAnnotationOptions(deps);

  const yomitanTokens = await parseWithYomitanInternalParser(tokenizeText, deps, annotationOptions);
  if (yomitanTokens && yomitanTokens.length > 0) {
    return {
      text: displayText,
      tokens: await applyAnnotationStage(yomitanTokens, deps, annotationOptions),
    };
  }

  return { text: displayText, tokens: null };
}
