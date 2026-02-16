import { BrowserWindow, Extension, session } from "electron";
import { markNPlusOneTargets, mergeTokens } from "../../token-merger";
import {
  JlptLevel,
  MergedToken,
  NPlusOneMatchMode,
  PartOfSpeech,
  SubtitleData,
  Token,
} from "../../types";
import { shouldIgnoreJlptForMecabPos1 } from "./jlpt-token-filter-config";
import { shouldIgnoreJlptByTerm } from "./jlpt-excluded-terms";

interface YomitanParseHeadword {
  term?: unknown;
}

interface YomitanParseSegment {
  text?: unknown;
  reading?: unknown;
  headwords?: unknown;
}

interface YomitanParseResultItem {
  source?: unknown;
  index?: unknown;
  content?: unknown;
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
    isKnownWord: options.isKnownWord,
    getKnownWordMatchMode: options.getKnownWordMatchMode,
    getJlptLevel: options.getJlptLevel,
    getJlptEnabled: options.getJlptEnabled,
    tokenizeWithMecab: async (text) => {
      const mecabTokenizer = options.getMecabTokenizer();
      if (!mecabTokenizer) {
        return null;
      }
      const rawTokens = await mecabTokenizer.tokenize(text);
      if (!rawTokens || rawTokens.length === 0) {
        return null;
      }
      return mergeTokens(
        rawTokens,
        options.isKnownWord,
        options.getKnownWordMatchMode(),
      );
    },
  };
}

function resolveKnownWordText(
  surface: string,
  headword: string,
  matchMode: NPlusOneMatchMode,
): string {
  return matchMode === "surface" ? surface : headword;
}

function applyKnownWordMarking(
  tokens: MergedToken[],
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): MergedToken[] {
  return tokens.map((token) => {
    const matchText = resolveKnownWordText(
      token.surface,
      token.headword,
      knownWordMatchMode,
    );

    return {
      ...token,
      isKnown: token.isKnown || (matchText ? isKnownWord(matchText) : false),
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
    return "";
  }

  let normalized = "";
  for (const char of raw) {
    const code = char.codePointAt(0);
    if (code === undefined) {
      continue;
    }

    if (code >= 0x30a1 && code <= 0x30f6) {
      normalized += String.fromCodePoint(code - 0x60);
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
  ].filter((candidate): candidate is string => typeof candidate === "string" && candidate.length > 0);

  for (const candidate of candidates) {
    const normalizedCandidate = normalizeJlptTextForExclusion(candidate);
    if (!normalizedCandidate) {
      continue;
    }

    const trimmedCandidate = candidate.trim();
    if (
      shouldIgnoreJlptByTerm(trimmedCandidate) ||
      shouldIgnoreJlptByTerm(normalizedCandidate)
    ) {
      return false;
    }

    if (
      isRepeatedKanaSfx(candidate) ||
      isRepeatedKanaSfx(normalizedCandidate)
    ) {
      return false;
    }
  }

  return true;
}

function applyJlptMarking(
  tokens: MergedToken[],
  getJlptLevel: (text: string) => JlptLevel | null,
): MergedToken[] {
  return tokens.map((token) => {
    if (!isJlptEligibleToken(token)) {
      return { ...token, jlptLevel: undefined };
    }

    const primaryLevel = getJlptLevel(resolveJlptLookupText(token));
    const fallbackLevel = getJlptLevel(token.surface);

    return {
      ...token,
      jlptLevel: primaryLevel ?? fallbackLevel ?? token.jlptLevel,
    };
  });
}

function extractYomitanHeadword(segment: YomitanParseSegment): string {
  const headwords = segment.headwords;
  if (!Array.isArray(headwords) || headwords.length === 0) {
    return "";
  }

  const firstGroup = headwords[0];
  if (!Array.isArray(firstGroup) || firstGroup.length === 0) {
    return "";
  }

  const firstHeadword = firstGroup[0] as YomitanParseHeadword;
  return typeof firstHeadword?.term === "string" ? firstHeadword.term : "";
}

function mapYomitanParseResultsToMergedTokens(
  parseResults: unknown,
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
  getJlptLevel: (text: string) => JlptLevel | null,
): MergedToken[] | null {
  if (!Array.isArray(parseResults) || parseResults.length === 0) {
    return null;
  }

  const scanningItems = parseResults.filter((item) => {
    const resultItem = item as YomitanParseResultItem;
    return (
      resultItem &&
      resultItem.source === "scanning-parser" &&
      Array.isArray(resultItem.content)
    );
  }) as YomitanParseResultItem[];

  if (scanningItems.length === 0) {
    return null;
  }

  const primaryItem =
    scanningItems.find((item) => item.index === 0) || scanningItems[0];
  const content = primaryItem.content;
  if (!Array.isArray(content)) {
    return null;
  }

  const tokens: MergedToken[] = [];
  let charOffset = 0;

  for (const line of content) {
    if (!Array.isArray(line)) {
      continue;
    }

    let surface = "";
    let reading = "";
    let headword = "";

    for (const rawSegment of line) {
      const segment = rawSegment as YomitanParseSegment;
      if (!segment || typeof segment !== "object") {
        continue;
      }

      const segmentText = segment.text;
      if (typeof segmentText !== "string" || segmentText.length === 0) {
        continue;
      }

      surface += segmentText;

      if (typeof segment.reading === "string") {
        reading += segment.reading;
      }

      if (!headword) {
        headword = extractYomitanHeadword(segment);
      }
    }

    if (!surface) {
      continue;
    }

    const start = charOffset;
    const end = start + surface.length;
    charOffset = end;

    tokens.push({
      surface,
      reading,
      headword: headword || surface,
      startPos: start,
      endPos: end,
      partOfSpeech: PartOfSpeech.other,
      pos1: "",
      isMerged: true,
      isNPlusOneTarget: false,
      isKnown: (() => {
        const matchText = resolveKnownWordText(
          surface,
          headword,
          knownWordMatchMode,
        );
        return matchText ? isKnownWord(matchText) : false;
      })(),
    });
  }

  return tokens.length > 0 ? tokens : null;
}

function pickClosestMecabPos1(
  token: MergedToken,
  mecabTokens: MergedToken[],
): string | undefined {
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
        (span > bestSpan ||
          (span === bestSpan && mecabStart < bestStart)))
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
    console.warn(
      "Failed to enrich Yomitan tokens with MeCab POS:",
      (err as Error).message,
    );
    return tokens;
  }

  if (!mecabTokens || mecabTokens.length === 0) {
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

async function ensureYomitanParserWindow(
  deps: TokenizerServiceDeps,
): Promise<boolean> {
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
        parserWindow.webContents.once("did-finish-load", () => resolve());
        parserWindow.webContents.once(
          "did-fail-load",
          (_event, _errorCode, errorDescription) => {
            reject(new Error(errorDescription));
          },
        );
      }),
    );

    parserWindow.on("closed", () => {
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
      console.error(
        "Failed to initialize Yomitan parser window:",
        (err as Error).message,
      );
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
        useMecabParser: false
      });
    })();
  `;

  try {
    const parseResults = await parserWindow.webContents.executeJavaScript(
      script,
      true,
    );
    const yomitanTokens = mapYomitanParseResultsToMergedTokens(
      parseResults,
      deps.isKnownWord,
      deps.getKnownWordMatchMode(),
      deps.getJlptLevel,
    );
    if (!yomitanTokens || yomitanTokens.length === 0) {
      return null;
    }

    return enrichYomitanPos1(yomitanTokens, deps, text);
  } catch (err) {
    console.error("Yomitan parser request failed:", (err as Error).message);
    return null;
  }
}

export async function tokenizeSubtitleService(
  text: string,
  deps: TokenizerServiceDeps,
): Promise<SubtitleData> {
  const displayText = text
    .replace(/\r\n/g, "\n")
    .replace(/\\N/g, "\n")
    .replace(/\\n/g, "\n")
    .trim();

  if (!displayText) {
    return { text, tokens: null };
  }

  const tokenizeText = displayText
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const jlptEnabled = deps.getJlptEnabled?.() !== false;

  const yomitanTokens = await parseWithYomitanInternalParser(tokenizeText, deps);
  if (yomitanTokens && yomitanTokens.length > 0) {
    const knownMarkedTokens = applyKnownWordMarking(
      yomitanTokens,
      deps.isKnownWord,
      deps.getKnownWordMatchMode(),
    );
    const jlptMarkedTokens = jlptEnabled
      ? applyJlptMarking(knownMarkedTokens, deps.getJlptLevel)
      : knownMarkedTokens.map((token) => ({ ...token, jlptLevel: undefined }));
    return { text: displayText, tokens: markNPlusOneTargets(jlptMarkedTokens) };
  }

  try {
    const mecabTokens = await deps.tokenizeWithMecab(tokenizeText);
    if (mecabTokens && mecabTokens.length > 0) {
      const knownMarkedTokens = applyKnownWordMarking(
        mecabTokens,
        deps.isKnownWord,
        deps.getKnownWordMatchMode(),
      );
      const jlptMarkedTokens = jlptEnabled
        ? applyJlptMarking(knownMarkedTokens, deps.getJlptLevel)
        : knownMarkedTokens.map((token) => ({ ...token, jlptLevel: undefined }));
      return { text: displayText, tokens: markNPlusOneTargets(jlptMarkedTokens) };
    }
  } catch (err) {
    console.error("Tokenization error:", (err as Error).message);
  }

  return { text: displayText, tokens: null };
}
