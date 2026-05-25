import type { BrowserWindow, Extension, Session } from 'electron';
import { mergeTokens } from '../../token-merger';
import { createLogger } from '../../logger';
import {
  FrequencyDictionaryMatchMode,
  CharacterNameImage,
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
import type { YomitanTermFrequency } from './tokenizer/yomitan-parser-runtime';

const logger = createLogger('main:tokenizer');

type MecabTokenEnrichmentFn = (
  tokens: MergedToken[],
  mecabTokens: MergedToken[] | null,
) => Promise<MergedToken[]>;

export interface TokenizerServiceDeps {
  getYomitanExt: () => Extension | null;
  getYomitanSession?: () => Session | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  isKnownWord: (text: string) => boolean;
  getKnownWordMatchMode: () => NPlusOneMatchMode;
  getKnownWordsEnabled?: () => boolean;
  getJlptLevel: (text: string) => JlptLevel | null;
  getNPlusOneEnabled?: () => boolean;
  getJlptEnabled?: () => boolean;
  getNameMatchEnabled?: () => boolean;
  getNameMatchImagesEnabled?: () => boolean;
  getCharacterNameImage?: (term: string) => CharacterNameImage | null;
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
  getYomitanSession?: () => Session | null;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => Promise<void> | null;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => Promise<boolean> | null;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  isKnownWord: (text: string) => boolean;
  getKnownWordMatchMode: () => NPlusOneMatchMode;
  getKnownWordsEnabled?: () => boolean;
  getJlptLevel: (text: string) => JlptLevel | null;
  getNPlusOneEnabled?: () => boolean;
  getJlptEnabled?: () => boolean;
  getNameMatchEnabled?: () => boolean;
  getNameMatchImagesEnabled?: () => boolean;
  getCharacterNameImage?: (term: string) => CharacterNameImage | null;
  getFrequencyDictionaryEnabled?: () => boolean;
  getFrequencyDictionaryMatchMode?: () => FrequencyDictionaryMatchMode;
  getFrequencyRank?: FrequencyDictionaryLookup;
  getMinSentenceWordsForNPlusOne?: () => number;
  getYomitanGroupDebugEnabled?: () => boolean;
  getMecabTokenizer: () => MecabTokenizerLike | null;
  onTokenizationReady?: (text: string) => void;
}

interface TokenizerAnnotationOptions {
  knownWordsEnabled: boolean;
  nPlusOneEnabled: boolean;
  jlptEnabled: boolean;
  nameMatchEnabled: boolean;
  nameMatchImagesEnabled: boolean;
  frequencyEnabled: boolean;
  frequencyMatchMode: FrequencyDictionaryMatchMode;
  minSentenceWordsForNPlusOne: number | undefined;
  pos1Exclusions: ReadonlySet<string>;
  pos2Exclusions: ReadonlySet<string>;
  sourceText?: string;
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
const INVISIBLE_SEPARATOR_PATTERN = /[\u200b\u2060\ufeff]/g;

function getKnownWordLookup(
  deps: TokenizerServiceDeps,
  options: TokenizerAnnotationOptions,
): (text: string) => boolean {
  if (!options.knownWordsEnabled && !options.nPlusOneEnabled) {
    return () => false;
  }
  return deps.isKnownWord;
}

function needsMecabPosEnrichment(options: TokenizerAnnotationOptions): boolean {
  return (
    options.knownWordsEnabled ||
    options.nPlusOneEnabled ||
    options.jlptEnabled ||
    options.frequencyEnabled
  );
}

function hasAnyAnnotationEnabled(options: TokenizerAnnotationOptions): boolean {
  return (
    options.knownWordsEnabled ||
    options.nPlusOneEnabled ||
    options.jlptEnabled ||
    options.frequencyEnabled
  );
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
    return stripSubtitleAnnotationMetadata(tokens, options);
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

async function stripSubtitleAnnotationMetadata(
  tokens: MergedToken[],
  options: TokenizerAnnotationOptions,
): Promise<MergedToken[]> {
  if (tokens.length === 0) {
    return tokens;
  }

  if (!annotationStageModulePromise) {
    annotationStageModulePromise = import('./tokenizer/annotation-stage');
  }

  const annotationStage = await annotationStageModulePromise;
  return tokens.map((token) => annotationStage.stripSubtitleAnnotationMetadata(token, options));
}

export function createTokenizerDepsRuntime(
  options: TokenizerDepsRuntimeOptions,
): TokenizerServiceDeps {
  const checkedMecabTokenizers = new WeakSet<object>();

  return {
    getYomitanExt: options.getYomitanExt,
    getYomitanSession: options.getYomitanSession,
    getYomitanParserWindow: options.getYomitanParserWindow,
    setYomitanParserWindow: options.setYomitanParserWindow,
    getYomitanParserReadyPromise: options.getYomitanParserReadyPromise,
    setYomitanParserReadyPromise: options.setYomitanParserReadyPromise,
    getYomitanParserInitPromise: options.getYomitanParserInitPromise,
    setYomitanParserInitPromise: options.setYomitanParserInitPromise,
    isKnownWord: options.isKnownWord,
    getKnownWordMatchMode: options.getKnownWordMatchMode,
    getKnownWordsEnabled: options.getKnownWordsEnabled,
    getJlptLevel: options.getJlptLevel,
    getNPlusOneEnabled: options.getNPlusOneEnabled,
    getJlptEnabled: options.getJlptEnabled,
    getNameMatchEnabled: options.getNameMatchEnabled,
    getNameMatchImagesEnabled: options.getNameMatchImagesEnabled,
    getCharacterNameImage: options.getCharacterNameImage,
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

      return mergeTokens(
        rawTokens,
        options.isKnownWord,
        options.getKnownWordMatchMode(),
        false,
        text,
      );
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
    isNameMatch: token.isNameMatch ?? false,
    reading: normalizeYomitanMergedReading(token),
  }));
}

function normalizeYomitanWordClasses(wordClasses: unknown): string[] {
  if (!Array.isArray(wordClasses)) {
    return [];
  }

  const normalized: string[] = [];
  for (const wordClass of wordClasses) {
    if (typeof wordClass !== 'string') {
      continue;
    }
    const trimmed = wordClass.trim();
    if (trimmed && !normalized.includes(trimmed)) {
      normalized.push(trimmed);
    }
  }
  return normalized;
}

function resolvePartOfSpeechFromYomitanWordClasses(wordClasses: string[]): {
  partOfSpeech: PartOfSpeech;
  pos1?: string;
} {
  if (wordClasses.includes('prt')) {
    return { partOfSpeech: PartOfSpeech.particle, pos1: '助詞' };
  }
  if (wordClasses.some((wordClass) => wordClass === 'aux' || wordClass.startsWith('aux-'))) {
    return { partOfSpeech: PartOfSpeech.bound_auxiliary, pos1: '助動詞' };
  }
  if (wordClasses.some((wordClass) => wordClass.startsWith('v'))) {
    return { partOfSpeech: PartOfSpeech.verb, pos1: '動詞' };
  }
  if (wordClasses.includes('adj-i') || wordClasses.includes('adj-ix')) {
    return { partOfSpeech: PartOfSpeech.i_adjective, pos1: '形容詞' };
  }
  if (wordClasses.includes('adj-na')) {
    return { partOfSpeech: PartOfSpeech.na_adjective, pos1: '名詞' };
  }
  if (
    wordClasses.some(
      (wordClass) =>
        wordClass === 'n' ||
        wordClass === 'num' ||
        wordClass === 'ctr' ||
        wordClass === 'pn' ||
        wordClass.startsWith('n-'),
    )
  ) {
    return { partOfSpeech: PartOfSpeech.noun, pos1: '名詞' };
  }

  return { partOfSpeech: PartOfSpeech.other };
}

function getYomitanWordClassPosMetadata(wordClasses: unknown): {
  partOfSpeech: PartOfSpeech;
  pos1?: string;
} {
  return resolvePartOfSpeechFromYomitanWordClasses(normalizeYomitanWordClasses(wordClasses));
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

function resolveYomitanFrequencyLookupTexts(
  token: MergedToken,
  matchMode: FrequencyDictionaryMatchMode,
): string[] {
  const primaryLookupText = resolveFrequencyLookupText(token, matchMode).trim();
  if (!primaryLookupText) {
    return [];
  }

  if (matchMode !== 'headword') {
    return [primaryLookupText];
  }

  const normalizedHeadword = token.headword.trim();
  const normalizedSurface = token.surface.trim();
  if (
    !normalizedHeadword ||
    !normalizedSurface ||
    normalizedSurface === normalizedHeadword ||
    normalizedSurface === primaryLookupText
  ) {
    return [primaryLookupText];
  }

  return [primaryLookupText, normalizedSurface];
}

function buildYomitanFrequencyTermReadingList(
  tokens: MergedToken[],
  matchMode: FrequencyDictionaryMatchMode,
): Array<{ term: string; reading: string | null }> {
  const termReadingList: Array<{ term: string; reading: string | null }> = [];
  for (const token of tokens) {
    const readingRaw =
      token.reading && token.reading.trim().length > 0 ? token.reading.trim() : null;
    for (const term of resolveYomitanFrequencyLookupTexts(token, matchMode)) {
      termReadingList.push({ term, reading: readingRaw });
    }
  }

  return termReadingList;
}

function makeYomitanFrequencyPairKey(term: string, reading: string | null): string {
  return `${term}\u0000${reading ?? ''}`;
}

interface NormalizedYomitanTermFrequency extends YomitanTermFrequency {
  reading: string | null;
  frequency: number;
}

interface YomitanFrequencyIndex {
  byPair: Map<string, NormalizedYomitanTermFrequency[]>;
  byTerm: Map<string, NormalizedYomitanTermFrequency[]>;
}

function appendYomitanFrequencyEntry(
  map: Map<string, NormalizedYomitanTermFrequency[]>,
  key: string,
  entry: NormalizedYomitanTermFrequency,
): void {
  const existing = map.get(key);
  if (existing) {
    existing.push(entry);
    return;
  }

  map.set(key, [entry]);
}

function buildYomitanFrequencyIndex(
  frequencies: ReadonlyArray<YomitanTermFrequency>,
): YomitanFrequencyIndex {
  const byPair = new Map<string, NormalizedYomitanTermFrequency[]>();
  const byTerm = new Map<string, NormalizedYomitanTermFrequency[]>();
  for (const frequency of frequencies) {
    const term = frequency.term.trim();
    const rank = normalizePositiveFrequencyRank(frequency.frequency);
    if (!term || rank === null) {
      continue;
    }

    const reading =
      typeof frequency.reading === 'string' && frequency.reading.trim().length > 0
        ? frequency.reading.trim()
        : null;
    const normalizedEntry: NormalizedYomitanTermFrequency = {
      ...frequency,
      term,
      reading,
      frequency: rank,
    };
    appendYomitanFrequencyEntry(
      byPair,
      makeYomitanFrequencyPairKey(term, reading),
      normalizedEntry,
    );
    appendYomitanFrequencyEntry(byTerm, term, normalizedEntry);
  }

  return { byPair, byTerm };
}

function selectBestYomitanFrequencyRank(
  entries: ReadonlyArray<NormalizedYomitanTermFrequency>,
): number | null {
  let bestEntry: NormalizedYomitanTermFrequency | null = null;
  for (const entry of entries) {
    if (
      bestEntry === null ||
      entry.dictionaryPriority < bestEntry.dictionaryPriority ||
      (entry.dictionaryPriority === bestEntry.dictionaryPriority &&
        entry.frequency < bestEntry.frequency)
    ) {
      bestEntry = entry;
    }
  }

  return bestEntry?.frequency ?? null;
}

function getYomitanFrequencyRank(
  token: MergedToken,
  candidateText: string,
  matchMode: FrequencyDictionaryMatchMode,
  frequencyIndex: YomitanFrequencyIndex,
): number | null {
  const normalizedCandidateText = candidateText.trim();
  if (!normalizedCandidateText) {
    return null;
  }

  const reading =
    typeof token.reading === 'string' && token.reading.trim().length > 0
      ? token.reading.trim()
      : null;
  const pairEntries =
    frequencyIndex.byPair.get(makeYomitanFrequencyPairKey(normalizedCandidateText, reading)) ?? [];
  const candidateEntries =
    pairEntries.length > 0
      ? pairEntries
      : (frequencyIndex.byTerm.get(normalizedCandidateText) ?? []);
  if (candidateEntries.length === 0) {
    return null;
  }

  const normalizedHeadword = token.headword.trim();
  const normalizedSurface = token.surface.trim();
  const isInflectedHeadwordFallback =
    matchMode === 'headword' &&
    normalizedCandidateText === normalizedHeadword &&
    normalizedSurface.length > 0 &&
    normalizedSurface !== normalizedHeadword;

  return selectBestYomitanFrequencyRank(candidateEntries);
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
  yomitanFrequencyIndex: YomitanFrequencyIndex,
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

    for (const candidateText of resolveYomitanFrequencyLookupTexts(token, matchMode)) {
      const yomitanRank = getYomitanFrequencyRank(
        token,
        candidateText,
        matchMode,
        yomitanFrequencyIndex,
      );
      if (yomitanRank !== null) {
        return {
          ...token,
          frequencyRank: yomitanRank,
        };
      }
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
  const nPlusOneEnabled = deps.getNPlusOneEnabled?.() !== false;
  return {
    knownWordsEnabled: deps.getKnownWordsEnabled
      ? deps.getKnownWordsEnabled() !== false
      : nPlusOneEnabled,
    nPlusOneEnabled,
    jlptEnabled: deps.getJlptEnabled?.() !== false,
    nameMatchEnabled: deps.getNameMatchEnabled?.() !== false,
    nameMatchImagesEnabled: deps.getNameMatchImagesEnabled?.() === true,
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
  const selectedTokens = await requestYomitanScanTokens(text, deps, logger, {
    includeNameMatchMetadata: options.nameMatchEnabled,
  });
  if (!selectedTokens || selectedTokens.length === 0) {
    return null;
  }
  const normalizedSelectedTokens = normalizeSelectedYomitanTokens(
    selectedTokens.map((token): MergedToken => {
      const posMetadata = getYomitanWordClassPosMetadata(token.wordClasses);
      return {
        surface: token.surface,
        reading: token.reading,
        headword: token.headword,
        startPos: token.startPos,
        endPos: token.endPos,
        partOfSpeech: posMetadata.partOfSpeech,
        pos1: posMetadata.pos1,
        isMerged: true,
        isKnown: false,
        isNPlusOneTarget: false,
        isNameMatch: token.isNameMatch ?? false,
        frequencyRank: token.frequencyRank,
      };
    }),
  );

  if (deps.getYomitanGroupDebugEnabled?.() === true) {
    logSelectedYomitanGroups(text, normalizedSelectedTokens);
  }
  deps.onTokenizationReady?.(text);

  const frequencyRankPromise: Promise<YomitanFrequencyIndex> = options.frequencyEnabled
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
        return buildYomitanFrequencyIndex(yomitanFrequencies);
      })()
    : Promise.resolve({ byPair: new Map(), byTerm: new Map() });

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

  const [yomitanFrequencyIndex, enrichedTokens] = await Promise.all([
    frequencyRankPromise,
    mecabEnrichmentPromise,
  ]);

  if (options.frequencyEnabled) {
    return applyFrequencyRanks(
      enrichedTokens,
      options.frequencyMatchMode,
      yomitanFrequencyIndex,
      deps.getFrequencyRank,
    );
  }

  return enrichedTokens;
}

function resolveCharacterNameImageForToken(
  token: MergedToken,
  getCharacterNameImage: (term: string) => CharacterNameImage | null,
): CharacterNameImage | null {
  const terms = [token.headword, token.surface]
    .map((term) => term.trim())
    .filter((term, index, list) => term.length > 0 && list.indexOf(term) === index);
  for (const term of terms) {
    const image = getCharacterNameImage(term);
    if (image) {
      return image;
    }
  }
  return null;
}

function applyCharacterNameImages(
  tokens: MergedToken[],
  deps: TokenizerServiceDeps,
  options: TokenizerAnnotationOptions,
): MergedToken[] {
  if (
    !options.nameMatchEnabled ||
    !options.nameMatchImagesEnabled ||
    typeof deps.getCharacterNameImage !== 'function'
  ) {
    return tokens.map((token) => ({ ...token, characterImage: undefined }));
  }

  return tokens.map((token) => {
    if (token.isNameMatch !== true) {
      return { ...token, characterImage: undefined };
    }
    return {
      ...token,
      characterImage:
        resolveCharacterNameImageForToken(token, deps.getCharacterNameImage!) ?? undefined,
    };
  });
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

  const tokenizeText = displayText
    .replace(INVISIBLE_SEPARATOR_PATTERN, ' ')
    .replace(/\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const annotationOptions = getAnnotationOptions(deps);
  annotationOptions.sourceText = tokenizeText;

  const yomitanTokens = await parseWithYomitanInternalParser(tokenizeText, deps, annotationOptions);
  if (yomitanTokens && yomitanTokens.length > 0) {
    const annotatedTokens = await applyAnnotationStage(yomitanTokens, deps, annotationOptions);
    const renderedTokens = applyCharacterNameImages(annotatedTokens, deps, annotationOptions);
    return {
      text: displayText,
      tokens: renderedTokens.length > 0 ? renderedTokens : null,
    };
  }

  return { text: displayText, tokens: null };
}
