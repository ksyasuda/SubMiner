import fs from "node:fs";
import path from "node:path";
import process from "node:process";

import { createTokenizerDepsRuntime, tokenizeSubtitle } from "../src/core/services/tokenizer.js";
import { createFrequencyDictionaryLookup } from "../src/core/services/frequency-dictionary.js";
import { MecabTokenizer } from "../src/mecab-tokenizer.js";
import type { MergedToken, FrequencyDictionaryLookup } from "../src/types.js";

interface CliOptions {
  input: string;
  dictionaryPath: string;
  emitPretty: boolean;
  emitDiagnostics: boolean;
  mecabCommand?: string;
  mecabDictionaryPath?: string;
  forceMecabOnly?: boolean;
  yomitanExtensionPath?: string;
  yomitanUserDataPath?: string;
  emitColoredLine: boolean;
  colorMode: "single" | "banded";
  colorTopX: number;
  colorSingle: string;
  colorBand1: string;
  colorBand2: string;
  colorBand3: string;
  colorBand4: string;
  colorBand5: string;
  colorKnown: string;
  colorNPlusOne: string;
}

function parseCliArgs(argv: string[]): CliOptions {
  const args = [...argv];
  let inputParts: string[] = [];
  let dictionaryPath = path.join(process.cwd(), "vendor", "jiten_freq_global");
  let emitPretty = false;
  let emitDiagnostics = false;
  let mecabCommand: string | undefined;
  let mecabDictionaryPath: string | undefined;
  let forceMecabOnly = false;
  let yomitanExtensionPath: string | undefined;
  let yomitanUserDataPath: string | undefined;
  let emitColoredLine = false;
  let colorMode: "single" | "banded" = "single";
  let colorTopX = 1000;
  let colorSingle = "#f5a97f";
  let colorBand1 = "#ed8796";
  let colorBand2 = "#f5a97f";
  let colorBand3 = "#f9e2af";
  let colorBand4 = "#a6e3a1";
  let colorBand5 = "#8aadf4";
  let colorKnown = "#a6da95";
  let colorNPlusOne = "#c6a0f6";

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;

    if (arg === "--help" || arg === "-h") {
      printUsage();
      process.exit(0);
    }

    if (arg === "--dictionary") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --dictionary");
      }
      dictionaryPath = path.resolve(next);
      continue;
    }

    if (arg === "--mecab-command") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --mecab-command");
      }
      mecabCommand = next;
      continue;
    }

    if (arg === "--mecab-dictionary") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --mecab-dictionary");
      }
      mecabDictionaryPath = next;
      continue;
    }

    if (arg === "--yomitan-extension") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --yomitan-extension");
      }
      yomitanExtensionPath = path.resolve(next);
      continue;
    }

    if (arg === "--yomitan-user-data") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --yomitan-user-data");
      }
      yomitanUserDataPath = path.resolve(next);
      continue;
    }

    if (arg === "--colorized-line") {
      emitColoredLine = true;
      continue;
    }

    if (arg === "--color-mode") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-mode");
      }
      if (next !== "single" && next !== "banded") {
        throw new Error("--color-mode must be 'single' or 'banded'");
      }
      colorMode = next;
      continue;
    }

    if (arg === "--color-top-x") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-top-x");
      }
      const parsed = Number.parseInt(next, 10);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        throw new Error("--color-top-x must be a positive integer");
      }
      colorTopX = parsed;
      continue;
    }

    if (arg === "--color-single") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-single");
      }
      colorSingle = next;
      continue;
    }

    if (arg === "--color-band-1") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-band-1");
      }
      colorBand1 = next;
      continue;
    }

    if (arg === "--color-band-2") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-band-2");
      }
      colorBand2 = next;
      continue;
    }

    if (arg === "--color-band-3") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-band-3");
      }
      colorBand3 = next;
      continue;
    }

    if (arg === "--color-band-4") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-band-4");
      }
      colorBand4 = next;
      continue;
    }

    if (arg === "--color-band-5") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-band-5");
      }
      colorBand5 = next;
      continue;
    }

    if (arg === "--color-known") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-known");
      }
      colorKnown = next;
      continue;
    }

    if (arg === "--color-n-plus-one") {
      const next = args.shift();
      if (!next) {
        throw new Error("Missing value for --color-n-plus-one");
      }
      colorNPlusOne = next;
      continue;
    }

    if (arg.startsWith("--dictionary=")) {
      dictionaryPath = path.resolve(arg.slice("--dictionary=".length));
      continue;
    }

    if (arg.startsWith("--mecab-command=")) {
      mecabCommand = arg.slice("--mecab-command=".length);
      continue;
    }

    if (arg.startsWith("--mecab-dictionary=")) {
      mecabDictionaryPath = arg.slice("--mecab-dictionary=".length);
      continue;
    }

    if (arg.startsWith("--yomitan-extension=")) {
      yomitanExtensionPath = path.resolve(
        arg.slice("--yomitan-extension=".length),
      );
      continue;
    }

    if (arg.startsWith("--yomitan-user-data=")) {
      yomitanUserDataPath = path.resolve(
        arg.slice("--yomitan-user-data=".length),
      );
      continue;
    }

    if (arg.startsWith("--colorized-line")) {
      emitColoredLine = true;
      continue;
    }

    if (arg.startsWith("--color-mode=")) {
      const value = arg.slice("--color-mode=".length);
      if (value !== "single" && value !== "banded") {
        throw new Error("--color-mode must be 'single' or 'banded'");
      }
      colorMode = value;
      continue;
    }

    if (arg.startsWith("--color-top-x=")) {
      const value = arg.slice("--color-top-x=".length);
      const parsed = Number.parseInt(value, 10);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        throw new Error("--color-top-x must be a positive integer");
      }
      colorTopX = parsed;
      continue;
    }

    if (arg.startsWith("--color-single=")) {
      colorSingle = arg.slice("--color-single=".length);
      continue;
    }

    if (arg.startsWith("--color-band-1=")) {
      colorBand1 = arg.slice("--color-band-1=".length);
      continue;
    }

    if (arg.startsWith("--color-band-2=")) {
      colorBand2 = arg.slice("--color-band-2=".length);
      continue;
    }

    if (arg.startsWith("--color-band-3=")) {
      colorBand3 = arg.slice("--color-band-3=".length);
      continue;
    }

    if (arg.startsWith("--color-band-4=")) {
      colorBand4 = arg.slice("--color-band-4=".length);
      continue;
    }

    if (arg.startsWith("--color-band-5=")) {
      colorBand5 = arg.slice("--color-band-5=".length);
      continue;
    }

    if (arg.startsWith("--color-known=")) {
      colorKnown = arg.slice("--color-known=".length);
      continue;
    }

    if (arg.startsWith("--color-n-plus-one=")) {
      colorNPlusOne = arg.slice("--color-n-plus-one=".length);
      continue;
    }

    if (arg === "--pretty") {
      emitPretty = true;
      continue;
    }

    if (arg === "--diagnostics") {
      emitDiagnostics = true;
      continue;
    }

    if (arg === "--force-mecab") {
      forceMecabOnly = true;
      continue;
    }

    if (arg.startsWith("-")) {
      throw new Error(`Unknown flag: ${arg}`);
    }

    inputParts.push(arg);
  }

  const input = inputParts.join(" ").trim();
  if (!input) {
    const stdin = fs.readFileSync(0, "utf8").trim();
    if (!stdin) {
      throw new Error(
        "Please provide input text as arguments or via stdin.",
      );
    }
    return {
      input: stdin,
      dictionaryPath,
      emitPretty,
      emitDiagnostics,
      forceMecabOnly,
      yomitanExtensionPath,
      yomitanUserDataPath,
      emitColoredLine,
      colorMode,
      colorTopX,
      colorSingle,
      colorBand1,
      colorBand2,
      colorBand3,
      colorBand4,
      colorBand5,
      colorKnown,
      colorNPlusOne,
      mecabCommand,
      mecabDictionaryPath,
    };
  }

  return {
    input,
    dictionaryPath,
    emitPretty,
    emitDiagnostics,
    forceMecabOnly,
    yomitanExtensionPath,
    yomitanUserDataPath,
    emitColoredLine,
    colorMode,
    colorTopX,
    colorSingle,
    colorBand1,
    colorBand2,
    colorBand3,
    colorBand4,
    colorBand5,
    colorKnown,
    colorNPlusOne,
    mecabCommand,
    mecabDictionaryPath,
  };
  }

function printUsage(): void {
  process.stdout.write(`Usage:
   pnpm run get-frequency [--pretty] [--diagnostics] [--dictionary <path>] [--mecab-command <path>] [--mecab-dictionary <path>] <text>

  --pretty               Pretty-print JSON output.
  --diagnostics            Include merged-frequency lookup-term details.
  --force-mecab          Skip Yomitan parser initialization and force MeCab fallback.
  --yomitan-extension <path> Optional path to a Yomitan extension directory.
  --yomitan-user-data <path> Optional Electron userData directory for Yomitan state.
  --colorized-line        Output a terminal-colorized line based on token classification.
  --color-mode <single|banded> Frequency coloring mode (default: single).
  --color-top-x <n>      Frequency color applies when rank <= n (default: 1000).
  --color-single <#hex>  Frequency single-mode color (default: #f5a97f).
  --color-band-1 <#hex>  Frequency band-1 color.
  --color-band-2 <#hex>  Frequency band-2 color.
  --color-band-3 <#hex>  Frequency band-3 color.
  --color-band-4 <#hex>  Frequency band-4 color.
  --color-band-5 <#hex>  Frequency band-5 color.
  --color-known <#hex>    Known-word color (default: #a6da95).
  --color-n-plus-one <#hex> N+1 target color (default: #c6a0f6).
  --dictionary <path>    Frequency dictionary root path (default: ./vendor/jiten_freq_global)
  --mecab-command <path>  Optional MeCab binary path (default: mecab)
  --mecab-dictionary <path> Optional MeCab dictionary directory (default: system default)
  -h, --help            Show usage.
\n`);
}

type FrequencyCandidate = {
  term: string;
  rank: number;
};

function getFrequencyLookupTextCandidates(token: MergedToken): string[] {
  const lookupText = token.headword?.trim() || token.reading?.trim() || token.surface.trim();
  return lookupText ? [lookupText] : [];
}

function getBestFrequencyLookupCandidate(
  token: MergedToken,
  getFrequencyRank: FrequencyDictionaryLookup,
): FrequencyCandidate | null {
  const lookupTexts = getFrequencyLookupTextCandidates(token);
  let best: FrequencyCandidate | null = null;
  for (const term of lookupTexts) {
    const rank = getFrequencyRank(term);
    if (typeof rank !== "number" || !Number.isFinite(rank) || rank <= 0) {
      continue;
    }
    if (!best || rank < best.rank) {
      best = { term, rank };
    }
  }
  return best;
}

function simplifyToken(token: MergedToken): Record<string, unknown> {
  return {
    surface: token.surface,
    reading: token.reading,
    headword: token.headword,
    startPos: token.startPos,
    endPos: token.endPos,
    partOfSpeech: token.partOfSpeech,
    isMerged: token.isMerged,
    isKnown: token.isKnown,
    isNPlusOneTarget: token.isNPlusOneTarget,
    frequencyRank: token.frequencyRank,
    jlptLevel: token.jlptLevel,
  };
}

function simplifyTokenWithVerbose(
  token: MergedToken,
  getFrequencyRank: FrequencyDictionaryLookup,
): Record<string, unknown> {
  const candidates = getFrequencyLookupTextCandidates(token).map((term) => ({
    term,
    rank: getFrequencyRank(term),
  })).filter((candidate) =>
    typeof candidate.rank === "number" &&
    Number.isFinite(candidate.rank) &&
    candidate.rank > 0
  );

  const bestCandidate = getBestFrequencyLookupCandidate(
    token,
    getFrequencyRank,
  );

  return {
    surface: token.surface,
    reading: token.reading,
    headword: token.headword,
    startPos: token.startPos,
    endPos: token.endPos,
    partOfSpeech: token.partOfSpeech,
    isMerged: token.isMerged,
    isKnown: token.isKnown,
    isNPlusOneTarget: token.isNPlusOneTarget,
    frequencyRank: token.frequencyRank,
    jlptLevel: token.jlptLevel,
    frequencyCandidates: candidates,
    frequencyBestLookupTerm: bestCandidate?.term ?? null,
    frequencyBestLookupRank: bestCandidate?.rank ?? null,
  };
}

interface YomitanRuntimeState {
  yomitanExt: unknown | null;
  parserWindow: unknown | null;
  parserReadyPromise: Promise<void> | null;
  parserInitPromise: Promise<boolean> | null;
  available: boolean;
  note?: string;
}

function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  label: string,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`${label} timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    promise
      .then((value) => {
        clearTimeout(timer);
        resolve(value);
      })
      .catch((error) => {
        clearTimeout(timer);
        reject(error);
      });
  });
}

function destroyUnknownParserWindow(window: unknown): void {
  if (!window || typeof window !== "object") {
    return;
  }
  const candidate = window as {
    isDestroyed?: () => boolean;
    destroy?: () => void;
  };
  if (typeof candidate.isDestroyed !== "function") {
    return;
  }
  if (typeof candidate.destroy !== "function") {
    return;
  }
  if (!candidate.isDestroyed()) {
    candidate.destroy();
  }
}

async function createYomitanRuntimeState(
  userDataPath: string,
): Promise<YomitanRuntimeState> {
  const state: YomitanRuntimeState = {
    yomitanExt: null,
    parserWindow: null,
    parserReadyPromise: null,
    parserInitPromise: null,
    available: false,
  };

  const electronImport = await import("electron").catch((error) => {
    state.note = error instanceof Error ? error.message : "unknown error";
    return null;
  });
  if (!electronImport || !electronImport.app || !electronImport.app.whenReady) {
    state.note = "electron runtime not available in this process";
    return state;
  }

  try {
    await electronImport.app.whenReady();
    const loadYomitanExtension = (
      await import(
        "../src/core/services/yomitan-extension-loader.js"
      )
    ).loadYomitanExtension as (
      options: {
        userDataPath: string;
        getYomitanParserWindow: () => unknown;
        setYomitanParserWindow: (window: unknown) => void;
        setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
        setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
        setYomitanExtension: (extension: unknown) => void;
      },
    ) => Promise<unknown>;

    const extension = await loadYomitanExtension({
      userDataPath,
      getYomitanParserWindow: () => state.parserWindow,
      setYomitanParserWindow: (window) => {
        state.parserWindow = window;
      },
      setYomitanParserReadyPromise: (promise) => {
        state.parserReadyPromise = promise;
      },
      setYomitanParserInitPromise: (promise) => {
        state.parserInitPromise = promise;
      },
      setYomitanExtension: (extension) => {
        state.yomitanExt = extension;
      },
    });

    if (!extension) {
      state.note = "yomitan extension is not available";
      return state;
    }

    state.yomitanExt = extension;
    state.available = true;
    return state;
  } catch (error) {
    state.note =
      error instanceof Error
        ? error.message
        : "failed to initialize yomitan extension";
    return state;
  }
}

async function createYomitanRuntimeStateWithSearch(
  userDataPath: string,
  extensionPath?: string,
): Promise<YomitanRuntimeState> {
  const preferredPath = extensionPath
    ? path.resolve(extensionPath)
    : undefined;
  const defaultVendorPath = path.resolve(process.cwd(), "vendor", "yomitan");
  const candidates = [
    ...(preferredPath ? [preferredPath] : []),
    defaultVendorPath,
  ];

  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }
    try {
      if (fs.existsSync(path.join(candidate, "manifest.json"))) {
        const state = await createYomitanRuntimeState(userDataPath);
        if (state.available) {
          return state;
        }
        if (!state.note) {
          state.note = `Failed to load yomitan extension at ${candidate}`;
        }
        return state;
      }
    } catch {
      continue;
    }
  }

  return createYomitanRuntimeState(userDataPath);
}

async function getFrequencyLookup(dictionaryPath: string): Promise<FrequencyDictionaryLookup> {
  return createFrequencyDictionaryLookup({
    searchPaths: [dictionaryPath],
    log: (message) => {
      // Keep script output pure JSON by default
      if (process.env.DEBUG_FREQUENCY === "1") {
        console.error(message);
      }
    },
  });
}

const ANSI_RESET = "\u001b[0m";
const ANSI_FG_PREFIX = "\u001b[38;2";
const HEX_COLOR_PATTERN = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;

function parseHexRgb(input: string): [number, number, number] | null {
  const normalized = input.trim().replace(/^#/, "");
  if (!HEX_COLOR_PATTERN.test(`#${normalized}`)) {
    return null;
  }
  const expanded = normalized.length === 3
    ? normalized.split("").map((char) => `${char}${char}`).join("")
    : normalized;
  const r = Number.parseInt(expanded.substring(0, 2), 16);
  const g = Number.parseInt(expanded.substring(2, 4), 16);
  const b = Number.parseInt(expanded.substring(4, 6), 16);
  if (
    !Number.isFinite(r) ||
    !Number.isFinite(g) ||
    !Number.isFinite(b)
  ) {
    return null;
  }
  return [r, g, b];
}

function wrapWithForeground(text: string, color: string): string {
  const rgb = parseHexRgb(color);
  if (!rgb) {
    return text;
  }
  return `${ANSI_FG_PREFIX};${rgb[0]};${rgb[1]};${rgb[2]}m${text}${ANSI_RESET}`;
}

function getBandColor(
  rank: number,
  colorTopX: number,
  colorMode: "single" | "banded",
  colorSingle: string,
  bandedColors: [string, string, string, string, string],
): string {
  const topX = Math.max(1, Math.floor(colorTopX));
  const safeRank = Math.max(1, Math.floor(rank));
  if (safeRank > topX) {
    return "";
  }
  if (colorMode === "single") {
    return colorSingle;
  }
  const normalizedBand = Math.ceil((safeRank / topX) * bandedColors.length);
  const band = Math.min(bandedColors.length, Math.max(1, normalizedBand));
  return bandedColors[band - 1];
}

function getTokenColor(token: MergedToken, args: CliOptions): string {
  if (token.isNPlusOneTarget) {
    return args.colorNPlusOne;
  }
  if (token.isKnown) {
    return args.colorKnown;
  }
  if (typeof token.frequencyRank === "number" && Number.isFinite(token.frequencyRank)) {
    return getBandColor(
      token.frequencyRank,
      args.colorTopX,
      args.colorMode,
      args.colorSingle,
      [args.colorBand1, args.colorBand2, args.colorBand3, args.colorBand4, args.colorBand5],
    );
  }
  return "";
}

function renderColoredLine(
  text: string,
  tokens: MergedToken[],
  args: CliOptions,
): string {
  if (!args.emitColoredLine) {
    return text;
  }
  if (tokens.length === 0) {
    return text;
  }

  const ordered = [...tokens].sort((a, b) => {
    const aStart = a.startPos ?? 0;
    const bStart = b.startPos ?? 0;
    if (aStart !== bStart) {
      return aStart - bStart;
    }
    return (a.endPos ?? a.surface.length) - (b.endPos ?? b.surface.length);
  });

  let cursor = 0;
  let output = "";
  for (const token of ordered) {
    const start = token.startPos ?? 0;
    const end = token.endPos ?? (token.startPos ? token.startPos + token.surface.length : token.surface.length);
    if (start < 0 || end < 0 || end < start) {
      continue;
    }
    const safeStart = Math.min(Math.max(0, start), text.length);
    const safeEnd = Math.min(Math.max(safeStart, end), text.length);
    if (safeStart > cursor) {
      output += text.slice(cursor, safeStart);
    }
    const tokenText = text.slice(safeStart, safeEnd);
    const color = getTokenColor(token, args);
    output += color ? wrapWithForeground(tokenText, color) : tokenText;
    cursor = safeEnd;
  }

  if (cursor < text.length) {
    output += text.slice(cursor);
  }
  return output;
}

async function main(): Promise<void> {
  let electronModule: (typeof import("electron")) | null = null;
  let yomitanState: YomitanRuntimeState | null = null;

  try {
    const args = parseCliArgs(process.argv.slice(2));
    const getFrequencyRank = await getFrequencyLookup(args.dictionaryPath);

    const mecabTokenizer = new MecabTokenizer({
      mecabCommand: args.mecabCommand,
      dictionaryPath: args.mecabDictionaryPath,
    });
    const isMecabAvailable = await mecabTokenizer.checkAvailability();
    if (!isMecabAvailable) {
      throw new Error(
        "MeCab is not available on this system. Install/run environment with MeCab to tokenize input.",
      );
    }

    electronModule = await import("electron").catch(() => null);
    if (electronModule && args.yomitanUserDataPath) {
      electronModule.app.setPath("userData", args.yomitanUserDataPath);
    }
    yomitanState =
      !args.forceMecabOnly
        ? await createYomitanRuntimeStateWithSearch(
            electronModule?.app?.getPath
              ? electronModule.app.getPath("userData")
              : process.cwd(),
            args.yomitanExtensionPath,
          )
        : null;
    const hasYomitan = Boolean(yomitanState?.available && yomitanState?.yomitanExt);
    let useYomitan = hasYomitan;

    const deps = createTokenizerDepsRuntime({
      getYomitanExt: () =>
        (useYomitan ? yomitanState!.yomitanExt : null) as never,
      getYomitanParserWindow: () =>
        (useYomitan ? yomitanState!.parserWindow : null) as never,
      setYomitanParserWindow: (window) => {
        if (!useYomitan) {
          return;
        }
        yomitanState!.parserWindow = window;
      },
      getYomitanParserReadyPromise: () =>
        (useYomitan ? yomitanState!.parserReadyPromise : null) as never,
      setYomitanParserReadyPromise: (promise) => {
        if (!useYomitan) {
          return;
        }
        yomitanState!.parserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () =>
        (useYomitan ? yomitanState!.parserInitPromise : null) as never,
      setYomitanParserInitPromise: (promise) => {
        if (!useYomitan) {
          return;
        }
        yomitanState!.parserInitPromise = promise;
      },
      isKnownWord: () => false,
      getKnownWordMatchMode: () => "headword",
      getJlptLevel: () => null,
      getFrequencyDictionaryEnabled: () => true,
      getFrequencyRank,
      getMecabTokenizer: () => ({
        tokenize: (text: string) => mecabTokenizer.tokenize(text),
      }),
    });

    let subtitleData;
    if (useYomitan) {
      try {
        subtitleData = await withTimeout(
          tokenizeSubtitle(args.input, deps),
          8000,
          "Yomitan tokenizer",
        );
      } catch (error) {
        useYomitan = false;
        destroyUnknownParserWindow(yomitanState?.parserWindow ?? null);
        if (yomitanState) {
          yomitanState.parserWindow = null;
          yomitanState.parserReadyPromise = null;
          yomitanState.parserInitPromise = null;
          const fallbackNote = error instanceof Error ? error.message : "Yomitan tokenizer timed out";
          yomitanState.note = yomitanState.note
            ? `${yomitanState.note}; ${fallbackNote}`
            : fallbackNote;
        }
        subtitleData = await tokenizeSubtitle(args.input, deps);
      }
    } else {
      subtitleData = await tokenizeSubtitle(args.input, deps);
    }
    const tokenCount = subtitleData.tokens?.length ?? 0;
    const mergedCount = subtitleData.tokens?.filter((token) => token.isMerged).length ?? 0;
    const tokens =
      subtitleData.tokens?.map((token) =>
        args.emitDiagnostics
          ? simplifyTokenWithVerbose(token, getFrequencyRank)
          : simplifyToken(token),
      ) ?? null;
    const diagnostics = {
      yomitan: {
        available: Boolean(yomitanState?.available),
        loaded: useYomitan,
        forceMecabOnly: args.forceMecabOnly,
        note: yomitanState?.note ?? null,
      },
      mecab: {
        command: args.mecabCommand ?? "mecab",
        dictionaryPath: args.mecabDictionaryPath ?? null,
        available: isMecabAvailable,
      },
      tokenizer: {
        sourceHint:
          tokenCount === 0
            ? "none"
            : useYomitan ? "yomitan-merged" : "mecab-merge",
        mergedTokenCount: mergedCount,
        totalTokenCount: tokenCount,
      },
    };
    if (tokens === null) {
      diagnostics.mecab["status"] = "no-tokens";
      diagnostics.mecab["note"] =
        "MeCab returned no parseable tokens. This is often caused by a missing/invalid MeCab dictionary path.";
    } else {
      diagnostics.mecab["status"] = "ok";
    }

    const output = {
      input: args.input,
      tokenizerText: subtitleData.text,
      tokens,
      diagnostics,
    };

    const json = JSON.stringify(output, null, args.emitPretty ? 2 : undefined);
    process.stdout.write(`${json}\n`);

    if (args.emitColoredLine && subtitleData.tokens) {
      const coloredLine = renderColoredLine(subtitleData.text, subtitleData.tokens, args);
      process.stdout.write(`${coloredLine}\n`);
    }
  } finally {
    destroyUnknownParserWindow(yomitanState?.parserWindow ?? null);
    if (electronModule?.app) {
      electronModule.app.quit();
    }
  }
}

main()
  .then(() => {
    process.exit(0);
  })
  .catch((error) => {
    console.error(`Error: ${(error as Error).message}`);
    process.exit(1);
  });
