import { BrowserWindow, Extension, session } from "electron";
import { markNPlusOneTargets, mergeTokens } from "../../token-merger";
import {
  MergedToken,
  NPlusOneMatchMode,
  PartOfSpeech,
  SubtitleData,
  Token,
} from "../../types";

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
    return mapYomitanParseResultsToMergedTokens(
      parseResults,
      deps.isKnownWord,
      deps.getKnownWordMatchMode(),
    );
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

  const yomitanTokens = await parseWithYomitanInternalParser(tokenizeText, deps);
  if (yomitanTokens && yomitanTokens.length > 0) {
    const knownMarkedTokens = applyKnownWordMarking(
      yomitanTokens,
      deps.isKnownWord,
      deps.getKnownWordMatchMode(),
    );
    return { text: displayText, tokens: markNPlusOneTargets(knownMarkedTokens) };
  }

  try {
    const mecabTokens = await deps.tokenizeWithMecab(tokenizeText);
    if (mecabTokens && mecabTokens.length > 0) {
      const knownMarkedTokens = applyKnownWordMarking(
        mecabTokens,
        deps.isKnownWord,
        deps.getKnownWordMatchMode(),
      );
      return { text: displayText, tokens: markNPlusOneTargets(knownMarkedTokens) };
    }
  } catch (err) {
    console.error("Tokenization error:", (err as Error).message);
  }

  return { text: displayText, tokens: null };
}
