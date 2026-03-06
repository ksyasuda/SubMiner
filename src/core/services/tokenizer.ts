import type { BrowserWindow, Extension } from 'electron';
import { mergeTokens } from '../../token-merger';
import { createLogger } from '../../logger';
import {
  FrequencyDictionaryMatchMode,
  MergedToken,
  NPlusOneMatchMode,
  SubtitleData,
  Token,
  FrequencyDictionaryLookup,
  JlptLevel,
  PartOfSpeech,
} from '../../types';
import {
  DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG,
  resolveAnnotationPos1ExclusionSet,
} from '../../token-pos1-exclusions';
import {
  DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG,
  resolveAnnotationPos2ExclusionSet,
} from '../../token-pos2-exclusions';
import {
  requestYomitanScanTokens,
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
  getFrequencyDictionaryMatchMode?: () => FrequencyDictionaryMatchMode;
  getFrequencyRank?: FrequencyDictionaryLookup;
  getMinSentenceWordsForNPlusOne?: () => number;
  getYomitanGroupDebugEnabled?: () => boolean;
  tokenizeWithMecab: (text: string) => Promise<MergedToken[] | null>;
  enrichTokensWithMecab?: MecabTokenEnrichmentFn;
  onTokenizationReady?: (text: string) => void;
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
  getFrequencyDictionaryMatchMode?: () => FrequencyDictionaryMatchMode;
  getFrequencyRank?: FrequencyDictionaryLookup;
  getMinSentenceWordsForNPlusOne?: () => number;
  getYomitanGroupDebugEnabled?: () => boolean;
  getMecabTokenizer: () => MecabTokenizerLike | null;
  onTokenizationReady?: (text: string) => void;
}

interface TokenizerAnnotationOptions {
  nPlusOneEnabled: boolean;
  jlptEnabled: boolean;
  frequencyEnabled: boolean;
  frequencyMatchMode: FrequencyDictionaryMatchMode;
  minSentenceWordsForNPlusOne: number | undefined;
  pos1Exclusions: ReadonlySet<string>;
  pos2Exclusions: ReadonlySet<string>;
}

let parserEnrichmentWorkerRuntimeModulePromise: Promise<
  typeof import('./tokenizer/parser-enrichment-worker-runtime')
> | null = null;
let annotationStageModulePromise: Promise<typeof import('./tokenizer/annotation-stage')> | null =
  null;
let parserEnrichmentFallbackModulePromise: Promise<
  typeof import('./tokenizer/parser-enrichment-stage')
> | null = null;
const DEFAULT_ANNOTATION_POS1_EXCLUSIONS = resolveAnnotationPos1ExclusionSet(
  DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG,
);
const DEFAULT_ANNOTATION_POS2_EXCLUSIONS = resolveAnnotationPos2ExclusionSet(
  DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG,
);

function getKnownWordLookup(
  deps: TokenizerServiceDeps,
  options: TokenizerAnnotationOptions,
): (text: string) => boolean {
  if (!options.nPlusOneEnabled) {
    return () => false;
  }
  return deps.isKnownWord;
}

function needsMecabPosEnrichment(options: TokenizerAnnotationOptions): boolean {
  return options.nPlusOneEnabled || options.jlptEnabled || options.frequencyEnabled;
}

function hasAnyAnnotationEnabled(options: TokenizerAnnotationOptions): boolean {
  return options.nPlusOneEnabled || options.jlptEnabled || options.frequencyEnabled;
}

async function enrichTokensWithMecabAsync(
  tokens: MergedToken[],
  mecabTokens: MergedToken[] | null,
): Promise<MergedToken[]> {
  if (!parserEnrichmentWorkerRuntimeModulePromise) {
    parserEnrichmentWorkerRuntimeModulePromise =
      import('./tokenizer/parser-enrichment-worker-runtime');
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
    getFrequencyDictionaryMatchMode: options.getFrequencyDictionaryMatchMode ?? (() => 'headword'),
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

      return mergeTokens(rawTokens, options.isKnownWord, options.getKnownWordMatchMode(), false);
    },
    enrichTokensWithMecab: async (tokens, mecabTokens) =>
      enrichTokensWithMecabAsync(tokens, mecabTokens),
    onTokenizationReady: options.onTokenizationReady,
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

function normalizeFrequencyLookupText(rawText: string): string {
  return rawText.trim().toLowerCase();
}

function isKanaChar(char: string): boolean {
  const code = char.codePointAt(0);
  if (code === undefined) {
    return false;
  }
  return (
    (code >= 0x3041 && code <= 0x3096) ||
    (code >= 0x309b && code <= 0x309f) ||
    code === 0x30fc ||
    (code >= 0x30a0 && code <= 0x30fa) ||
    (code >= 0x30fd && code <= 0x30ff)
  );
}

function getTrailingKanaSuffix(surface: string): string {
  const chars = Array.from(surface);
  let splitIndex = chars.length;
  while (splitIndex > 0 && isKanaChar(chars[splitIndex - 1]!)) {
    splitIndex -= 1;
  }
  if (splitIndex <= 0 || splitIndex >= chars.length) {
    return '';
  }
  return chars.slice(splitIndex).join('');
}

function normalizeYomitanMergedReading(token: MergedToken): string {
  const reading = token.reading ?? '';
  if (!reading || token.headword !== token.surface) {
    return reading;
  }
  const trailingKanaSuffix = getTrailingKanaSuffix(token.surface);
  if (!trailingKanaSuffix || reading.endsWith(trailingKanaSuffix)) {
    return reading;
  }
  return `${reading}${trailingKanaSuffix}`;
}

function normalizeSelectedYomitanTokens(tokens: MergedToken[]): MergedToken[] {
  return tokens.map((token) => ({
    ...token,
    partOfSpeech: token.partOfSpeech ?? PartOfSpeech.other,
    isMerged: token.isMerged ?? true,
    isKnown: token.isKnown ?? false,
    isNPlusOneTarget: token.isNPlusOneTarget ?? false,
    reading: normalizeYomitanMergedReading(token),
  }));
}

function resolveFrequencyLookupText(
  token: MergedToken,
  matchMode: FrequencyDictionaryMatchMode,
): string {
  if (matchMode === 'surface') {
    if (token.surface && token.surface.length > 0) {
      return token.surface;
    }
    if (token.headword && token.headword.length > 0) {
      return token.headword;
    }
    return token.reading;
  }

  if (token.headword && token.headword.length > 0) {
    return token.headword;
  }
  if (token.reading && token.reading.length > 0) {
    return token.reading;
  }
  return token.surface;
}

function buildYomitanFrequencyTermReadingList(
  tokens: MergedToken[],
  matchMode: FrequencyDictionaryMatchMode,
): Array<{ term: string; reading: string | null }> {
  const termReadingList: Array<{ term: string; reading: string | null }> = [];
  for (const token of tokens) {
    const term = resolveFrequencyLookupText(token, matchMode).trim();
    if (!term) {
      continue;
    }

    const readingRaw =
      token.reading && token.reading.trim().length > 0 ? token.reading.trim() : null;
    termReadingList.push({ term, reading: readingRaw });
  }

  return termReadingList;
}

function buildYomitanFrequencyRankMap(
  frequencies: ReadonlyArray<{ term: string; frequency: number; dictionaryPriority?: number }>,
): Map<string, number> {
  const rankByTerm = new Map<string, { rank: number; dictionaryPriority: number }>();
  for (const frequency of frequencies) {
    const normalizedTerm = frequency.term.trim();
    const rank = normalizePositiveFrequencyRank(frequency.frequency);
    if (!normalizedTerm || rank === null) {
      continue;
    }
    const dictionaryPriority =
      typeof frequency.dictionaryPriority === 'number' &&
      Number.isFinite(frequency.dictionaryPriority)
        ? Math.max(0, Math.floor(frequency.dictionaryPriority))
        : Number.MAX_SAFE_INTEGER;
    const current = rankByTerm.get(normalizedTerm);
    if (
      current === undefined ||
      dictionaryPriority < current.dictionaryPriority ||
      (dictionaryPriority === current.dictionaryPriority && rank < current.rank)
    ) {
      rankByTerm.set(normalizedTerm, { rank, dictionaryPriority });
    }
  }

  const collapsedRankByTerm = new Map<string, number>();
  for (const [term, entry] of rankByTerm.entries()) {
    collapsedRankByTerm.set(term, entry.rank);
  }

  return collapsedRankByTerm;
}

function getLocalFrequencyRank(
  lookupText: string,
  getFrequencyRank: FrequencyDictionaryLookup,
  cache: Map<string, number | null>,
): number | null {
  const normalizedText = normalizeFrequencyLookupText(lookupText);
  if (!normalizedText) {
    return null;
  }

  if (cache.has(normalizedText)) {
    return cache.get(normalizedText) ?? null;
  }

  let rank: number | null;
  try {
    rank = getFrequencyRank(normalizedText);
  } catch {
    rank = null;
  }
  rank = normalizePositiveFrequencyRank(rank);
  cache.set(normalizedText, rank);
  return rank;
}

function applyFrequencyRanks(
  tokens: MergedToken[],
  matchMode: FrequencyDictionaryMatchMode,
  yomitanRankByTerm: Map<string, number>,
  getFrequencyRank: FrequencyDictionaryLookup | undefined,
): MergedToken[] {
  if (tokens.length === 0) {
    return tokens;
  }

  const localLookupCache = new Map<string, number | null>();
  return tokens.map((token) => {
    const existingRank = normalizePositiveFrequencyRank(token.frequencyRank);
    if (existingRank !== null) {
      return {
        ...token,
        frequencyRank: existingRank,
      };
    }

    const lookupText = resolveFrequencyLookupText(token, matchMode).trim();
    if (!lookupText) {
      return {
        ...token,
        frequencyRank: undefined,
      };
    }

    const yomitanRank = yomitanRankByTerm.get(lookupText);
    if (yomitanRank !== undefined) {
      return {
        ...token,
        frequencyRank: yomitanRank,
      };
    }

    if (!getFrequencyRank) {
      return {
        ...token,
        frequencyRank: undefined,
      };
    }

    const localRank = getLocalFrequencyRank(lookupText, getFrequencyRank, localLookupCache);
    return {
      ...token,
      frequencyRank: localRank ?? undefined,
    };
  });
}

function getAnnotationOptions(deps: TokenizerServiceDeps): TokenizerAnnotationOptions {
  return {
    nPlusOneEnabled: deps.getNPlusOneEnabled?.() !== false,
    jlptEnabled: deps.getJlptEnabled?.() !== false,
    frequencyEnabled: deps.getFrequencyDictionaryEnabled?.() !== false,
    frequencyMatchMode: deps.getFrequencyDictionaryMatchMode?.() ?? 'headword',
    minSentenceWordsForNPlusOne: deps.getMinSentenceWordsForNPlusOne?.(),
    pos1Exclusions: DEFAULT_ANNOTATION_POS1_EXCLUSIONS,
    pos2Exclusions: DEFAULT_ANNOTATION_POS2_EXCLUSIONS,
  };
}

async function parseWithYomitanInternalParser(
  text: string,
  deps: TokenizerServiceDeps,
  options: TokenizerAnnotationOptions,
): Promise<MergedToken[] | null> {
  const selectedTokens = await requestYomitanScanTokens(text, deps, logger);
  if (!selectedTokens || selectedTokens.length === 0) {
    return null;
  }
  const normalizedSelectedTokens = normalizeSelectedYomitanTokens(
    selectedTokens.map(
      (token): MergedToken => ({
        surface: token.surface,
        reading: token.reading,
        headword: token.headword,
        startPos: token.startPos,
        endPos: token.endPos,
        partOfSpeech: PartOfSpeech.other,
        isMerged: true,
        isKnown: false,
        isNPlusOneTarget: false,
      }),
    ),
  );

  if (deps.getYomitanGroupDebugEnabled?.() === true) {
    logSelectedYomitanGroups(text, normalizedSelectedTokens);
  }
  deps.onTokenizationReady?.(text);

  const frequencyRankPromise: Promise<Map<string, number>> = options.frequencyEnabled
    ? (async () => {
        const frequencyMatchMode = options.frequencyMatchMode;
        const termReadingList = buildYomitanFrequencyTermReadingList(
          normalizedSelectedTokens,
          frequencyMatchMode,
        );
        const yomitanFrequencies = await requestYomitanTermFrequencies(
          termReadingList,
          deps,
          logger,
        );
        return buildYomitanFrequencyRankMap(yomitanFrequencies);
      })()
    : Promise.resolve(new Map<string, number>());

  const mecabEnrichmentPromise: Promise<MergedToken[]> = needsMecabPosEnrichment(options)
    ? (async () => {
        try {
          const mecabTokens = await deps.tokenizeWithMecab(text);
          const enrichTokensWithMecab = deps.enrichTokensWithMecab ?? enrichTokensWithMecabAsync;
          return await enrichTokensWithMecab(normalizedSelectedTokens, mecabTokens);
        } catch (err) {
          const error = err as Error;
          logger.warn(
            'Failed to enrich Yomitan tokens with MeCab POS:',
            error.message,
            `tokenCount=${normalizedSelectedTokens.length}`,
            `textLength=${text.length}`,
          );
          return normalizedSelectedTokens;
        }
      })()
    : Promise.resolve(normalizedSelectedTokens);

  const [yomitanRankByTerm, enrichedTokens] = await Promise.all([
    frequencyRankPromise,
    mecabEnrichmentPromise,
  ]);

  if (options.frequencyEnabled) {
    return applyFrequencyRanks(
      enrichedTokens,
      options.frequencyMatchMode,
      yomitanRankByTerm,
      deps.getFrequencyRank,
    );
  }

  return enrichedTokens;
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
