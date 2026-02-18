import type { BrowserWindow, Extension } from 'electron';
import { markNPlusOneTargets, mergeTokens } from '../../token-merger';
import {
  JlptLevel,
  MergedToken,
  NPlusOneMatchMode,
  PartOfSpeech,
  SubtitleData,
  Token,
  FrequencyDictionaryLookup,
} from '../../types';
import { shouldIgnoreJlptForMecabPos1, shouldIgnoreJlptByTerm } from './jlpt-token-filter';
import { createLogger } from '../../logger';

interface YomitanParseHeadword {
  term?: unknown;
}

interface YomitanParseSegment {
  text?: string;
  reading?: string;
  headwords?: unknown;
}

interface YomitanParseResultItem {
  source?: unknown;
  index?: unknown;
  content?: unknown;
}

type YomitanParseLine = YomitanParseSegment[];

const KATAKANA_TO_HIRAGANA_OFFSET = 0x60;
const KATAKANA_CODEPOINT_START = 0x30a1;
const KATAKANA_CODEPOINT_END = 0x30f6;
const JLPT_LEVEL_LOOKUP_CACHE_LIMIT = 2048;
const FREQUENCY_RANK_LOOKUP_CACHE_LIMIT = 2048;
const logger = createLogger('main:tokenizer');

const jlptLevelLookupCaches = new WeakMap<
  (text: string) => JlptLevel | null,
  Map<string, JlptLevel | null>
>();
const frequencyRankLookupCaches = new WeakMap<
  FrequencyDictionaryLookup,
  Map<string, number | null>
>();

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object');
}

function isString(value: unknown): value is string {
  return typeof value === 'string';
}

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

function getCachedJlptLevel(
  lookupText: string,
  getJlptLevel: (text: string) => JlptLevel | null,
): JlptLevel | null {
  const normalizedText = lookupText.trim();
  if (!normalizedText) {
    return null;
  }

  let cache = jlptLevelLookupCaches.get(getJlptLevel);
  if (!cache) {
    cache = new Map<string, JlptLevel | null>();
    jlptLevelLookupCaches.set(getJlptLevel, cache);
  }

  if (cache.has(normalizedText)) {
    return cache.get(normalizedText) ?? null;
  }

  let level: JlptLevel | null;
  try {
    level = getJlptLevel(normalizedText);
  } catch {
    level = null;
  }

  cache.set(normalizedText, level);
  while (cache.size > JLPT_LEVEL_LOOKUP_CACHE_LIMIT) {
    const firstKey = cache.keys().next().value;
    if (firstKey !== undefined) {
      cache.delete(firstKey);
    }
  }

  return level;
}

function normalizeFrequencyLookupText(rawText: string): string {
  return rawText.trim().toLowerCase();
}

function getCachedFrequencyRank(
  lookupText: string,
  getFrequencyRank: FrequencyDictionaryLookup,
): number | null {
  const normalizedText = normalizeFrequencyLookupText(lookupText);
  if (!normalizedText) {
    return null;
  }

  let cache = frequencyRankLookupCaches.get(getFrequencyRank);
  if (!cache) {
    cache = new Map<string, number | null>();
    frequencyRankLookupCaches.set(getFrequencyRank, cache);
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
  if (rank !== null) {
    if (!Number.isFinite(rank) || rank <= 0) {
      rank = null;
    }
  }

  cache.set(normalizedText, rank);
  while (cache.size > FREQUENCY_RANK_LOOKUP_CACHE_LIMIT) {
    const firstKey = cache.keys().next().value;
    if (firstKey !== undefined) {
      cache.delete(firstKey);
    }
  }

  return rank;
}

export function createTokenizerDepsRuntime(
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
      const rawTokens = await mecabTokenizer.tokenize(text);
      if (!rawTokens || rawTokens.length === 0) {
        return null;
      }
      return mergeTokens(rawTokens, options.isKnownWord, options.getKnownWordMatchMode());
    },
  };
}

function resolveKnownWordText(
  surface: string,
  headword: string,
  matchMode: NPlusOneMatchMode,
): string {
  return matchMode === 'surface' ? surface : headword;
}

function applyKnownWordMarking(
  tokens: MergedToken[],
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): MergedToken[] {
  return tokens.map((token) => {
    const matchText = resolveKnownWordText(token.surface, token.headword, knownWordMatchMode);

    return {
      ...token,
      isKnown: token.isKnown || (matchText ? isKnownWord(matchText) : false),
    };
  });
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

function getFrequencyLookupTextCandidates(token: MergedToken): string[] {
  const lookupText = resolveFrequencyLookupText(token).trim();
  return lookupText ? [lookupText] : [];
}

function isFrequencyExcludedByPos(token: MergedToken): boolean {
  if (
    token.partOfSpeech === PartOfSpeech.particle ||
    token.partOfSpeech === PartOfSpeech.bound_auxiliary
  ) {
    return true;
  }

  return token.pos1 === '助詞' || token.pos1 === '助動詞';
}

function applyFrequencyMarking(
  tokens: MergedToken[],
  getFrequencyRank: FrequencyDictionaryLookup,
): MergedToken[] {
  return tokens.map((token) => {
    if (isFrequencyExcludedByPos(token)) {
      return { ...token, frequencyRank: undefined };
    }

    const lookupTexts = getFrequencyLookupTextCandidates(token);
    if (lookupTexts.length === 0) {
      return { ...token, frequencyRank: undefined };
    }

    let bestRank: number | null = null;
    for (const lookupText of lookupTexts) {
      const rank = getCachedFrequencyRank(lookupText, getFrequencyRank);
      if (rank === null) {
        continue;
      }
      if (bestRank === null || rank < bestRank) {
        bestRank = rank;
      }
    }

    return {
      ...token,
      frequencyRank: bestRank ?? undefined,
    };
  });
}

function resolveJlptLookupText(token: MergedToken): string {
  if (token.headword && token.headword.length > 0) {
    return token.headword;
  }
  if (token.reading && token.reading.length > 0) {
    return token.reading;
  }
  return token.surface;
}

function normalizeJlptTextForExclusion(text: string): string {
  const raw = text.trim();
  if (!raw) {
    return '';
  }

  let normalized = '';
  for (const char of raw) {
    const code = char.codePointAt(0);
    if (code === undefined) {
      continue;
    }

    if (code >= KATAKANA_CODEPOINT_START && code <= KATAKANA_CODEPOINT_END) {
      normalized += String.fromCodePoint(code - KATAKANA_TO_HIRAGANA_OFFSET);
      continue;
    }

    normalized += char;
  }

  return normalized;
}

function isKanaChar(char: string): boolean {
  const code = char.codePointAt(0);
  if (code === undefined) {
    return false;
  }

  return (
    (code >= 0x3041 && code <= 0x3096) ||
    (code >= 0x309b && code <= 0x309f) ||
    (code >= 0x30a0 && code <= 0x30fa) ||
    (code >= 0x30fd && code <= 0x30ff)
  );
}

/**
 * Detects repeated-kana speech-like tokens (e.g. 「ああああ」, 「ははは」, 「うーん」 style patterns)
 * so they are not JLPT-labeled when they are mostly expressive particles/sfx.
 */
function isRepeatedKanaSfx(text: string): boolean {
  const normalized = text.trim();
  if (!normalized) {
    return false;
  }

  const chars = [...normalized];
  if (!chars.every(isKanaChar)) {
    return false;
  }

  const counts = new Map<string, number>();
  let hasAdjacentRepeat = false;

  for (let i = 0; i < chars.length; i += 1) {
    const char = chars[i];
    counts.set(char, (counts.get(char) ?? 0) + 1);
    if (i > 0 && chars[i] === chars[i - 1]) {
      hasAdjacentRepeat = true;
    }
  }

  const topCount = Math.max(...counts.values());
  if (chars.length <= 2) {
    return hasAdjacentRepeat || topCount >= 2;
  }

  if (hasAdjacentRepeat) {
    return true;
  }

  return topCount >= Math.ceil(chars.length / 2);
}

function isJlptEligibleToken(token: MergedToken): boolean {
  if (token.pos1 && shouldIgnoreJlptForMecabPos1(token.pos1)) return false;

  const candidates = [
    resolveJlptLookupText(token),
    token.surface,
    token.reading,
    token.headword,
  ].filter(
    (candidate): candidate is string => typeof candidate === 'string' && candidate.length > 0,
  );

  for (const candidate of candidates) {
    const normalizedCandidate = normalizeJlptTextForExclusion(candidate);
    if (!normalizedCandidate) {
      continue;
    }

    const trimmedCandidate = candidate.trim();
    if (shouldIgnoreJlptByTerm(trimmedCandidate) || shouldIgnoreJlptByTerm(normalizedCandidate)) {
      return false;
    }

    if (isRepeatedKanaSfx(candidate) || isRepeatedKanaSfx(normalizedCandidate)) {
      return false;
    }
  }

  return true;
}

function isYomitanParseResultItem(value: unknown): value is YomitanParseResultItem {
  if (!isObject(value)) {
    return false;
  }
  if (!isString((value as YomitanParseResultItem).source)) {
    return false;
  }
  if (!Array.isArray((value as YomitanParseResultItem).content)) {
    return false;
  }
  return true;
}

function isYomitanParseLine(value: unknown): value is YomitanParseLine {
  if (!Array.isArray(value)) {
    return false;
  }

  return value.every((segment) => {
    if (!isObject(segment)) {
      return false;
    }

    const candidate = segment as YomitanParseSegment;
    return isString(candidate.text);
  });
}

function isYomitanHeadwordRows(value: unknown): value is YomitanParseHeadword[][] {
  return (
    Array.isArray(value) &&
    value.every(
      (group) =>
        Array.isArray(group) &&
        group.every((item) => isObject(item) && isString((item as YomitanParseHeadword).term)),
    )
  );
}

function extractYomitanHeadword(segment: YomitanParseSegment): string {
  const headwords = segment.headwords;
  if (!isYomitanHeadwordRows(headwords)) {
    return '';
  }

  for (const group of headwords) {
    if (group.length > 0) {
      const firstHeadword = group[0] as YomitanParseHeadword;
      if (isString(firstHeadword?.term)) {
        return firstHeadword.term;
      }
    }
  }

  return '';
}

function applyJlptMarking(
  tokens: MergedToken[],
  getJlptLevel: (text: string) => JlptLevel | null,
): MergedToken[] {
  return tokens.map((token) => {
    if (!isJlptEligibleToken(token)) {
      return { ...token, jlptLevel: undefined };
    }

    const primaryLevel = getCachedJlptLevel(resolveJlptLookupText(token), getJlptLevel);
    const fallbackLevel =
      primaryLevel === null ? getCachedJlptLevel(token.surface, getJlptLevel) : null;

    return {
      ...token,
      jlptLevel: primaryLevel ?? fallbackLevel ?? token.jlptLevel,
    };
  });
}

interface YomitanParseCandidate {
  source: string;
  index: number;
  tokens: MergedToken[];
}

function mapYomitanParseResultItemToMergedTokens(
  parseResult: YomitanParseResultItem,
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): YomitanParseCandidate | null {
  const content = parseResult.content;
  if (!Array.isArray(content) || content.length === 0) {
    return null;
  }

  const source = String(parseResult.source ?? '');
  const index =
    typeof parseResult.index === 'number' && Number.isInteger(parseResult.index)
      ? parseResult.index
      : 0;

  const tokens: MergedToken[] = [];
  let charOffset = 0;
  let validLineCount = 0;

  for (const line of content) {
    if (!isYomitanParseLine(line)) {
      continue;
    }
    validLineCount += 1;

    let combinedSurface = '';
    let combinedReading = '';
    let combinedHeadword = '';

    for (const segment of line) {
      const segmentText = segment.text;
      if (!segmentText || segmentText.length === 0) {
        continue;
      }

      combinedSurface += segmentText;
      if (typeof segment.reading === 'string') {
        combinedReading += segment.reading;
      }
      if (!combinedHeadword) {
        combinedHeadword = extractYomitanHeadword(segment);
      }
    }

    if (!combinedSurface) {
      continue;
    }

    const start = charOffset;
    const end = start + combinedSurface.length;
    charOffset = end;
    const headword = combinedHeadword || combinedSurface;

    tokens.push({
      surface: combinedSurface,
      reading: combinedReading,
      headword,
      startPos: start,
      endPos: end,
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      isMerged: true,
      isNPlusOneTarget: false,
      isKnown: (() => {
        const matchText = resolveKnownWordText(combinedSurface, headword, knownWordMatchMode);
        return matchText ? isKnownWord(matchText) : false;
      })(),
    });
  }

  if (validLineCount === 0 || tokens.length === 0) {
    return null;
  }

  return { source, index, tokens };
}

function selectBestYomitanParseCandidate(
  candidates: YomitanParseCandidate[],
): MergedToken[] | null {
  if (candidates.length === 0) {
    return null;
  }

  const scanningCandidates = candidates.filter(
    (candidate) => candidate.source === 'scanning-parser',
  );
  const mecabCandidates = candidates.filter((candidate) => candidate.source === 'mecab');

  const getBestByTokenCount = (items: YomitanParseCandidate[]): YomitanParseCandidate | null =>
    items.length === 0
      ? null
      : items.reduce((best, current) =>
          current.tokens.length > best.tokens.length ? current : best,
        );

  const getCandidateScore = (candidate: YomitanParseCandidate): number => {
    const readableTokenCount = candidate.tokens.filter(
      (token) => token.reading.trim().length > 0,
    ).length;
    const suspiciousKanaFragmentCount = candidate.tokens.filter(
      (token) =>
        token.reading.trim().length === 0 &&
        token.surface.length >= 2 &&
        Array.from(token.surface).every((char) => isKanaChar(char)),
    ).length;

    return readableTokenCount * 100 - suspiciousKanaFragmentCount * 50 - candidate.tokens.length;
  };

  const chooseBestCandidate = (items: YomitanParseCandidate[]): YomitanParseCandidate | null => {
    if (items.length === 0) {
      return null;
    }

    return items.reduce((best, current) => {
      const bestScore = getCandidateScore(best);
      const currentScore = getCandidateScore(current);
      if (currentScore !== bestScore) {
        return currentScore > bestScore ? current : best;
      }

      if (current.tokens.length !== best.tokens.length) {
        return current.tokens.length < best.tokens.length ? current : best;
      }

      return best;
    });
  };

  if (scanningCandidates.length > 0) {
    const bestScanning = getBestByTokenCount(scanningCandidates);
    if (bestScanning && bestScanning.tokens.length > 1) {
      return bestScanning.tokens;
    }

    const bestMecab = chooseBestCandidate(mecabCandidates);
    if (bestMecab && bestMecab.tokens.length > (bestScanning?.tokens.length ?? 0)) {
      return bestMecab.tokens;
    }

    return bestScanning ? bestScanning.tokens : null;
  }

  const multiTokenCandidates = candidates.filter((candidate) => candidate.tokens.length > 1);
  const pool = multiTokenCandidates.length > 0 ? multiTokenCandidates : candidates;
  const bestCandidate = chooseBestCandidate(pool);
  return bestCandidate ? bestCandidate.tokens : null;
}

function mapYomitanParseResultsToMergedTokens(
  parseResults: unknown,
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): MergedToken[] | null {
  if (!Array.isArray(parseResults) || parseResults.length === 0) {
    return null;
  }

  const candidates = parseResults
    .filter((item): item is YomitanParseResultItem => isYomitanParseResultItem(item))
    .map((item) => mapYomitanParseResultItemToMergedTokens(item, isKnownWord, knownWordMatchMode))
    .filter((candidate): candidate is YomitanParseCandidate => candidate !== null);

  const bestCandidate = selectBestYomitanParseCandidate(candidates);
  return bestCandidate;
}

function logSelectedYomitanGroups(text: string, tokens: MergedToken[]): void {
  if (!tokens || tokens.length === 0) {
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

function pickClosestMecabPos1(token: MergedToken, mecabTokens: MergedToken[]): string | undefined {
  if (mecabTokens.length === 0) {
    return undefined;
  }

  const tokenStart = token.startPos ?? 0;
  const tokenEnd = token.endPos ?? tokenStart + token.surface.length;

  let bestPos1: string | undefined;
  let bestOverlap = 0;
  let bestSpan = 0;
  let bestStart = Number.MAX_SAFE_INTEGER;

  for (const mecabToken of mecabTokens) {
    if (!mecabToken.pos1) {
      continue;
    }

    const mecabStart = mecabToken.startPos ?? 0;
    const mecabEnd = mecabToken.endPos ?? mecabStart + mecabToken.surface.length;
    const overlapStart = Math.max(tokenStart, mecabStart);
    const overlapEnd = Math.min(tokenEnd, mecabEnd);
    const overlap = Math.max(0, overlapEnd - overlapStart);
    if (overlap === 0) {
      continue;
    }

    const span = mecabEnd - mecabStart;
    if (
      overlap > bestOverlap ||
      (overlap === bestOverlap &&
        (span > bestSpan || (span === bestSpan && mecabStart < bestStart)))
    ) {
      bestOverlap = overlap;
      bestSpan = span;
      bestStart = mecabStart;
      bestPos1 = mecabToken.pos1;
    }
  }

  return bestOverlap > 0 ? bestPos1 : undefined;
}

async function enrichYomitanPos1(
  tokens: MergedToken[],
  deps: TokenizerServiceDeps,
  text: string,
): Promise<MergedToken[]> {
  if (!tokens || tokens.length === 0) {
    return tokens;
  }

  let mecabTokens: MergedToken[] | null = null;
  try {
    mecabTokens = await deps.tokenizeWithMecab(text);
  } catch (err) {
    const error = err as Error;
    logger.warn(
      'Failed to enrich Yomitan tokens with MeCab POS:',
      error.message,
      `tokenCount=${tokens.length}`,
      `textLength=${text.length}`,
    );
    return tokens;
  }

  if (!mecabTokens || mecabTokens.length === 0) {
    logger.warn(
      'MeCab enrichment returned no tokens; preserving Yomitan token output.',
      `tokenCount=${tokens.length}`,
      `textLength=${text.length}`,
    );
    return tokens;
  }

  return tokens.map((token) => {
    if (token.pos1) {
      return token;
    }

    const pos1 = pickClosestMecabPos1(token, mecabTokens);
    if (!pos1) {
      return token;
    }

    return {
      ...token,
      pos1,
    };
  });
}

async function ensureYomitanParserWindow(deps: TokenizerServiceDeps): Promise<boolean> {
  const electron = await import('electron');
  const yomitanExt = deps.getYomitanExt();
  if (!yomitanExt) {
    return false;
  }

  const currentWindow = deps.getYomitanParserWindow();
  if (currentWindow && !currentWindow.isDestroyed()) {
    return true;
  }

  const existingInitPromise = deps.getYomitanParserInitPromise();
  if (existingInitPromise) {
    return existingInitPromise;
  }

  const initPromise = (async () => {
    const { BrowserWindow, session } = electron;
    const parserWindow = new BrowserWindow({
      show: false,
      width: 800,
      height: 600,
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false,
        session: session.defaultSession,
      },
    });
    deps.setYomitanParserWindow(parserWindow);

    deps.setYomitanParserReadyPromise(
      new Promise((resolve, reject) => {
        parserWindow.webContents.once('did-finish-load', () => resolve());
        parserWindow.webContents.once('did-fail-load', (_event, _errorCode, errorDescription) => {
          reject(new Error(errorDescription));
        });
      }),
    );

    parserWindow.on('closed', () => {
      if (deps.getYomitanParserWindow() === parserWindow) {
        deps.setYomitanParserWindow(null);
        deps.setYomitanParserReadyPromise(null);
      }
    });

    try {
      await parserWindow.loadURL(`chrome-extension://${yomitanExt.id}/search.html`);
      const readyPromise = deps.getYomitanParserReadyPromise();
      if (readyPromise) {
        await readyPromise;
      }
      return true;
    } catch (err) {
      logger.error('Failed to initialize Yomitan parser window:', (err as Error).message);
      if (!parserWindow.isDestroyed()) {
        parserWindow.destroy();
      }
      if (deps.getYomitanParserWindow() === parserWindow) {
        deps.setYomitanParserWindow(null);
        deps.setYomitanParserReadyPromise(null);
      }
      return false;
    } finally {
      deps.setYomitanParserInitPromise(null);
    }
  })();

  deps.setYomitanParserInitPromise(initPromise);
  return initPromise;
}

async function parseWithYomitanInternalParser(
  text: string,
  deps: TokenizerServiceDeps,
): Promise<MergedToken[] | null> {
  const yomitanExt = deps.getYomitanExt();
  if (!text || !yomitanExt) {
    return null;
  }

  const isReady = await ensureYomitanParserWindow(deps);
  const parserWindow = deps.getYomitanParserWindow();
  if (!isReady || !parserWindow || parserWindow.isDestroyed()) {
    return null;
  }

  const script = `
    (async () => {
      const invoke = (action, params) =>
        new Promise((resolve, reject) => {
          chrome.runtime.sendMessage({ action, params }, (response) => {
            if (chrome.runtime.lastError) {
              reject(new Error(chrome.runtime.lastError.message));
              return;
            }
            if (!response || typeof response !== "object") {
              reject(new Error("Invalid response from Yomitan backend"));
              return;
            }
            if (response.error) {
              reject(new Error(response.error.message || "Yomitan backend error"));
              return;
            }
            resolve(response.result);
          });
        });

      const optionsFull = await invoke("optionsGetFull", undefined);
      const profileIndex = optionsFull.profileCurrent;
      const scanLength =
        optionsFull.profiles?.[profileIndex]?.options?.scanning?.length ?? 40;

      return await invoke("parseText", {
        text: ${JSON.stringify(text)},
        optionsContext: { index: profileIndex },
        scanLength,
        useInternalParser: true,
        useMecabParser: true
      });
    })();
  `;

  try {
    const parseResults = await parserWindow.webContents.executeJavaScript(script, true);
    const yomitanTokens = mapYomitanParseResultsToMergedTokens(
      parseResults,
      deps.isKnownWord,
      deps.getKnownWordMatchMode(),
    );
    if (!yomitanTokens || yomitanTokens.length === 0) {
      return null;
    }

    if (deps.getYomitanGroupDebugEnabled?.() === true) {
      logSelectedYomitanGroups(text, yomitanTokens);
    }

    return enrichYomitanPos1(yomitanTokens, deps, text);
  } catch (err) {
    logger.error('Yomitan parser request failed:', (err as Error).message);
    return null;
  }
}

export async function tokenizeSubtitle(
  text: string,
  deps: TokenizerServiceDeps,
): Promise<SubtitleData> {
  const minSentenceWordsForNPlusOne = deps.getMinSentenceWordsForNPlusOne?.();
  const sanitizedMinSentenceWordsForNPlusOne =
    minSentenceWordsForNPlusOne !== undefined &&
    Number.isInteger(minSentenceWordsForNPlusOne) &&
    minSentenceWordsForNPlusOne > 0
      ? minSentenceWordsForNPlusOne
      : 3;

  const displayText = text
    .replace(/\r\n/g, '\n')
    .replace(/\\N/g, '\n')
    .replace(/\\n/g, '\n')
    .trim();

  if (!displayText) {
    return { text, tokens: null };
  }

  const tokenizeText = displayText.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
  const jlptEnabled = deps.getJlptEnabled?.() !== false;
  const frequencyEnabled = deps.getFrequencyDictionaryEnabled?.() !== false;
  const frequencyLookup = deps.getFrequencyRank;

  const yomitanTokens = await parseWithYomitanInternalParser(tokenizeText, deps);
  if (yomitanTokens && yomitanTokens.length > 0) {
    const knownMarkedTokens = applyKnownWordMarking(
      yomitanTokens,
      deps.isKnownWord,
      deps.getKnownWordMatchMode(),
    );
    const frequencyMarkedTokens =
      frequencyEnabled && frequencyLookup
        ? applyFrequencyMarking(knownMarkedTokens, frequencyLookup)
        : knownMarkedTokens.map((token) => ({
            ...token,
            frequencyRank: undefined,
          }));
    const jlptMarkedTokens = jlptEnabled
      ? applyJlptMarking(frequencyMarkedTokens, deps.getJlptLevel)
      : frequencyMarkedTokens.map((token) => ({
          ...token,
          jlptLevel: undefined,
        }));
    return {
      text: displayText,
      tokens: markNPlusOneTargets(jlptMarkedTokens, sanitizedMinSentenceWordsForNPlusOne),
    };
  }

  try {
    const mecabTokens = await deps.tokenizeWithMecab(tokenizeText);
    if (mecabTokens && mecabTokens.length > 0) {
      const knownMarkedTokens = applyKnownWordMarking(
        mecabTokens,
        deps.isKnownWord,
        deps.getKnownWordMatchMode(),
      );
      const frequencyMarkedTokens =
        frequencyEnabled && frequencyLookup
          ? applyFrequencyMarking(knownMarkedTokens, frequencyLookup)
          : knownMarkedTokens.map((token) => ({
              ...token,
              frequencyRank: undefined,
            }));
      const jlptMarkedTokens = jlptEnabled
        ? applyJlptMarking(frequencyMarkedTokens, deps.getJlptLevel)
        : frequencyMarkedTokens.map((token) => ({
            ...token,
            jlptLevel: undefined,
          }));
      return {
        text: displayText,
        tokens: markNPlusOneTargets(jlptMarkedTokens, sanitizedMinSentenceWordsForNPlusOne),
      };
    }
  } catch (err) {
    logger.error('Tokenization error:', (err as Error).message);
  }

  return { text: displayText, tokens: null };
}
