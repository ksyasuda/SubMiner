import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';

import { createTokenizerDepsRuntime, tokenizeSubtitle } from '../src/core/services/tokenizer.js';
import { resolveYomitanExtensionPath as resolveBuiltYomitanExtensionPath } from '../src/core/services/yomitan-extension-paths.js';
import { MecabTokenizer } from '../src/mecab-tokenizer.js';
import type { MergedToken } from '../src/types.js';

interface CliOptions {
  input: string;
  emitPretty: boolean;
  emitJson: boolean;
  forceMecabOnly: boolean;
  yomitanExtensionPath?: string;
  yomitanUserDataPath?: string;
  mecabCommand?: string;
  mecabDictionaryPath?: string;
}

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

interface ParsedCandidate {
  source: string;
  index: number;
  tokens: Array<{
    surface: string;
    reading: string;
    headword: string;
    startPos: number;
    endPos: number;
  }>;
}

interface YomitanRuntimeState {
  available: boolean;
  note: string | null;
  extension: Electron.Extension | null;
  parserWindow: Electron.BrowserWindow | null;
  parserReadyPromise: Promise<void> | null;
  parserInitPromise: Promise<boolean> | null;
}

const DEFAULT_YOMITAN_USER_DATA_PATH = path.join(os.homedir(), '.config', 'SubMiner');

function destroyParserWindow(window: Electron.BrowserWindow | null): void {
  if (!window || window.isDestroyed()) {
    return;
  }
  window.destroy();
}

async function shutdownYomitanRuntime(yomitan: YomitanRuntimeState): Promise<void> {
  destroyParserWindow(yomitan.parserWindow);
  const electronModule = await import('electron').catch(() => null);
  if (electronModule?.app) {
    electronModule.app.quit();
  }
}

function parseCliArgs(argv: string[]): CliOptions {
  const args = [...argv];
  const inputParts: string[] = [];
  let emitPretty = true;
  let emitJson = false;
  let forceMecabOnly = false;
  let yomitanExtensionPath: string | undefined;
  let yomitanUserDataPath: string | undefined = DEFAULT_YOMITAN_USER_DATA_PATH;
  let mecabCommand: string | undefined;
  let mecabDictionaryPath: string | undefined;

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;

    if (arg === '--help' || arg === '-h') {
      printUsage();
      process.exit(0);
    }

    if (arg === '--pretty') {
      emitPretty = true;
      continue;
    }

    if (arg === '--json') {
      emitJson = true;
      continue;
    }

    if (arg === '--force-mecab') {
      forceMecabOnly = true;
      continue;
    }

    if (arg === '--yomitan-extension') {
      const next = args.shift();
      if (!next) {
        throw new Error('Missing value for --yomitan-extension');
      }
      yomitanExtensionPath = next;
      continue;
    }

    if (arg.startsWith('--yomitan-extension=')) {
      yomitanExtensionPath = arg.slice('--yomitan-extension='.length);
      continue;
    }

    if (arg === '--yomitan-user-data') {
      const next = args.shift();
      if (!next) {
        throw new Error('Missing value for --yomitan-user-data');
      }
      yomitanUserDataPath = next;
      continue;
    }

    if (arg.startsWith('--yomitan-user-data=')) {
      yomitanUserDataPath = arg.slice('--yomitan-user-data='.length);
      continue;
    }

    if (arg === '--mecab-command') {
      const next = args.shift();
      if (!next) {
        throw new Error('Missing value for --mecab-command');
      }
      mecabCommand = next;
      continue;
    }

    if (arg.startsWith('--mecab-command=')) {
      mecabCommand = arg.slice('--mecab-command='.length);
      continue;
    }

    if (arg === '--mecab-dictionary') {
      const next = args.shift();
      if (!next) {
        throw new Error('Missing value for --mecab-dictionary');
      }
      mecabDictionaryPath = next;
      continue;
    }

    if (arg.startsWith('--mecab-dictionary=')) {
      mecabDictionaryPath = arg.slice('--mecab-dictionary='.length);
      continue;
    }

    if (arg.startsWith('-')) {
      throw new Error(`Unknown flag: ${arg}`);
    }

    inputParts.push(arg);
  }

  const input = inputParts.join(' ').trim();
  if (input.length > 0) {
    return {
      input,
      emitPretty,
      emitJson,
      forceMecabOnly,
      yomitanExtensionPath,
      yomitanUserDataPath,
      mecabCommand,
      mecabDictionaryPath,
    };
  }

  const stdin = fs.readFileSync(0, 'utf8').trim();
  if (!stdin) {
    throw new Error('Please provide input text as arguments or via stdin.');
  }

  return {
    input: stdin,
    emitPretty,
    emitJson,
    forceMecabOnly,
    yomitanExtensionPath,
    yomitanUserDataPath,
    mecabCommand,
    mecabDictionaryPath,
  };
}

function printUsage(): void {
  process.stdout.write(`Usage:
  bun run test-yomitan-parser:electron -- [--pretty] [--json] [--yomitan-extension <path>] [--yomitan-user-data <path>] [--mecab-command <path>] [--mecab-dictionary <path>] <text>

  --pretty               Pretty-print JSON output.
  --json                 Emit machine-readable JSON output.
  --force-mecab          Skip Yomitan parser setup and test MeCab fallback only.
  --yomitan-extension <path> Optional path to Yomitan extension directory.
  --yomitan-user-data <path> Optional Electron userData directory (default: ~/.config/SubMiner).
  --mecab-command <path> Optional MeCab binary path (default: mecab).
  --mecab-dictionary <path> Optional MeCab dictionary directory.
  -h, --help             Show usage.
`);
}

function normalizeDisplayText(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/\\N/g, '\n').replace(/\\n/g, '\n').trim();
}

function normalizeTokenizerText(text: string): string {
  return normalizeDisplayText(text).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
}

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object');
}

function isHeadwordRows(value: unknown): value is YomitanParseHeadword[][] {
  return (
    Array.isArray(value) &&
    value.every(
      (row) =>
        Array.isArray(row) &&
        row.every((entry) => isObject(entry) && typeof entry.term === 'string'),
    )
  );
}

function extractHeadwordTerms(segment: YomitanParseSegment): string[] {
  if (!isHeadwordRows(segment.headwords)) {
    return [];
  }
  const terms: string[] = [];
  const seen = new Set<string>();
  for (const row of segment.headwords) {
    for (const entry of row) {
      const term = (entry.term as string).trim();
      if (!term || seen.has(term)) {
        continue;
      }
      seen.add(term);
      terms.push(term);
    }
  }
  return terms;
}

function mapParseResultsToCandidates(parseResults: unknown): ParsedCandidate[] {
  if (!Array.isArray(parseResults)) {
    return [];
  }

  const candidates: ParsedCandidate[] = [];
  for (const item of parseResults) {
    if (!isObject(item)) {
      continue;
    }
    const parseItem = item as YomitanParseResultItem;
    if (!Array.isArray(parseItem.content) || typeof parseItem.source !== 'string') {
      continue;
    }

    const candidateTokens: ParsedCandidate['tokens'] = [];
    let charOffset = 0;
    let validLineCount = 0;

    for (const line of parseItem.content) {
      if (!Array.isArray(line)) {
        continue;
      }
      const lineSegments = line as YomitanParseSegment[];
      if (lineSegments.some((segment) => typeof segment.text !== 'string')) {
        continue;
      }
      validLineCount += 1;

      for (const segment of lineSegments) {
        const surface = (segment.text as string) ?? '';
        if (!surface) {
          continue;
        }
        const startPos = charOffset;
        const endPos = startPos + surface.length;
        charOffset = endPos;
        const headwordTerms = extractHeadwordTerms(segment);
        candidateTokens.push({
          surface,
          reading: typeof segment.reading === 'string' ? segment.reading : '',
          headword: headwordTerms[0] ?? surface,
          startPos,
          endPos,
        });
      }
    }

    if (validLineCount === 0 || candidateTokens.length === 0) {
      continue;
    }

    candidates.push({
      source: parseItem.source,
      index:
        typeof parseItem.index === 'number' && Number.isInteger(parseItem.index)
          ? parseItem.index
          : 0,
      tokens: candidateTokens,
    });
  }

  return candidates;
}

function candidateTokenSignature(token: {
  surface: string;
  reading: string;
  headword: string;
  startPos: number;
  endPos: number;
}): string {
  return `${token.surface}\u001f${token.reading}\u001f${token.headword}\u001f${token.startPos}\u001f${token.endPos}`;
}

function mergedTokenSignature(token: MergedToken): string {
  return `${token.surface}\u001f${token.reading}\u001f${token.headword}\u001f${token.startPos}\u001f${token.endPos}`;
}

function findSelectedCandidateIndexes(
  candidates: ParsedCandidate[],
  mergedTokens: MergedToken[] | null,
): number[] {
  if (!mergedTokens || mergedTokens.length === 0) {
    return [];
  }

  const mergedSignatures = mergedTokens.map(mergedTokenSignature);
  const selected: number[] = [];
  for (let i = 0; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    if (!candidate) {
      continue;
    }
    const candidateSignatures = candidate.tokens.map(candidateTokenSignature);
    if (candidateSignatures.length !== mergedSignatures.length) {
      continue;
    }
    let allMatch = true;
    for (let j = 0; j < candidateSignatures.length; j += 1) {
      if (candidateSignatures[j] !== mergedSignatures[j]) {
        allMatch = false;
        break;
      }
    }
    if (allMatch) {
      selected.push(i);
    }
  }

  return selected;
}

function resolveYomitanExtensionPath(explicitPath?: string): string | null {
  return resolveBuiltYomitanExtensionPath({
    explicitPath,
    cwd: process.cwd(),
  });
}

async function setupYomitanRuntime(options: CliOptions): Promise<YomitanRuntimeState> {
  const state: YomitanRuntimeState = {
    available: false,
    note: null,
    extension: null,
    parserWindow: null,
    parserReadyPromise: null,
    parserInitPromise: null,
  };

  if (options.forceMecabOnly) {
    state.note = 'force-mecab enabled';
    return state;
  }

  const electronModule = await import('electron').catch((error) => {
    state.note = error instanceof Error ? error.message : 'electron import failed';
    return null;
  });
  if (!electronModule?.app || !electronModule?.session) {
    state.note = 'electron runtime not available in this process';
    return state;
  }

  if (options.yomitanUserDataPath) {
    electronModule.app.setPath('userData', options.yomitanUserDataPath);
  }
  await electronModule.app.whenReady();

  const extensionPath = resolveYomitanExtensionPath(options.yomitanExtensionPath);
  if (!extensionPath) {
    state.note = 'no built Yomitan extension directory found; run `bun run build:yomitan`';
    return state;
  }

  try {
    state.extension = await electronModule.session.defaultSession.loadExtension(extensionPath, {
      allowFileAccess: true,
    });
    state.available = true;
    return state;
  } catch (error) {
    state.note = error instanceof Error ? error.message : 'failed to load Yomitan extension';
    state.available = false;
    return state;
  }
}

async function fetchRawParseResults(
  parserWindow: Electron.BrowserWindow,
  text: string,
): Promise<unknown> {
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
  return parserWindow.webContents.executeJavaScript(script, true);
}

function renderTextOutput(payload: Record<string, unknown>): void {
  process.stdout.write(`Input: ${String(payload.input)}\n`);
  process.stdout.write(`Tokenizer text: ${String(payload.tokenizerText)}\n`);
  process.stdout.write(`Yomitan available: ${String(payload.yomitanAvailable)}\n`);
  process.stdout.write(`Yomitan note: ${String(payload.yomitanNote ?? '')}\n`);
  process.stdout.write(
    `Selected candidate indexes: ${JSON.stringify(payload.selectedCandidateIndexes)}\n`,
  );
  process.stdout.write('\nFinal selected tokens:\n');
  const finalTokens = payload.finalTokens as Array<Record<string, unknown>> | null;
  if (!finalTokens || finalTokens.length === 0) {
    process.stdout.write('  (none)\n');
  } else {
    for (let i = 0; i < finalTokens.length; i += 1) {
      const token = finalTokens[i];
      if (!token) {
        continue;
      }
      process.stdout.write(
        `  [${i}] ${token.surface} -> ${token.headword} (${token.reading}) [${token.startPos}, ${token.endPos})\n`,
      );
    }
  }

  process.stdout.write('\nYomitan parse candidates:\n');
  const candidates = payload.candidates as Array<Record<string, unknown>>;
  if (!candidates || candidates.length === 0) {
    process.stdout.write('  (none)\n');
    return;
  }

  for (let i = 0; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    if (!candidate) {
      continue;
    }
    process.stdout.write(
      `  [${i}] source=${String(candidate.source)} index=${String(candidate.index)} selectedByTokenizer=${String(candidate.selectedByTokenizer)} tokenCount=${String(candidate.tokenCount)}\n`,
    );
    const tokens = candidate.tokens as Array<Record<string, unknown>> | undefined;
    if (!tokens || tokens.length === 0) {
      continue;
    }
    for (let j = 0; j < tokens.length; j += 1) {
      const token = tokens[j];
      if (!token) {
        continue;
      }
      process.stdout.write(
        `      - ${token.surface} -> ${token.headword} (${token.reading}) [${token.startPos}, ${token.endPos})\n`,
      );
    }
  }
}

async function main(): Promise<void> {
  const args = parseCliArgs(process.argv.slice(2));
  const yomitan: YomitanRuntimeState = {
    available: false,
    note: null,
    extension: null,
    parserWindow: null,
    parserReadyPromise: null,
    parserInitPromise: null,
  };

  try {
    const mecabTokenizer = new MecabTokenizer({
      mecabCommand: args.mecabCommand,
      dictionaryPath: args.mecabDictionaryPath,
    });
    const isMecabAvailable = await mecabTokenizer.checkAvailability();
    if (!isMecabAvailable) {
      throw new Error('MeCab is not available on this system.');
    }

    const runtime = await setupYomitanRuntime(args);
    yomitan.available = runtime.available;
    yomitan.note = runtime.note;
    yomitan.extension = runtime.extension;
    yomitan.parserWindow = runtime.parserWindow;
    yomitan.parserReadyPromise = runtime.parserReadyPromise;
    yomitan.parserInitPromise = runtime.parserInitPromise;

    const deps = createTokenizerDepsRuntime({
      getYomitanExt: () => yomitan.extension,
      getYomitanParserWindow: () => yomitan.parserWindow,
      setYomitanParserWindow: (window) => {
        yomitan.parserWindow = window;
      },
      getYomitanParserReadyPromise: () => yomitan.parserReadyPromise,
      setYomitanParserReadyPromise: (promise) => {
        yomitan.parserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () => yomitan.parserInitPromise,
      setYomitanParserInitPromise: (promise) => {
        yomitan.parserInitPromise = promise;
      },
      isKnownWord: () => false,
      getKnownWordMatchMode: () => 'headword',
      getJlptLevel: () => null,
      getMecabTokenizer: () => ({
        tokenize: (text: string) => mecabTokenizer.tokenize(text),
      }),
    });

    const subtitleData = await tokenizeSubtitle(args.input, deps);
    const tokenizeText = normalizeTokenizerText(args.input);
    let rawParseResults: unknown = null;
    if (
      yomitan.available &&
      yomitan.parserWindow &&
      !yomitan.parserWindow.isDestroyed() &&
      tokenizeText
    ) {
      rawParseResults = await fetchRawParseResults(yomitan.parserWindow, tokenizeText);
    }

    const parsedCandidates = mapParseResultsToCandidates(rawParseResults);
    const selectedCandidateIndexes = findSelectedCandidateIndexes(
      parsedCandidates,
      subtitleData.tokens,
    );
    const selectedIndexSet = new Set<number>(selectedCandidateIndexes);

    const payload = {
      input: args.input,
      tokenizerText: subtitleData.text,
      yomitanAvailable: yomitan.available,
      yomitanNote: yomitan.note,
      selectedCandidateIndexes,
      finalTokens:
        subtitleData.tokens?.map((token) => ({
          surface: token.surface,
          reading: token.reading,
          headword: token.headword,
          startPos: token.startPos,
          endPos: token.endPos,
          pos1: token.pos1,
          partOfSpeech: token.partOfSpeech,
          isKnown: token.isKnown,
          isNPlusOneTarget: token.isNPlusOneTarget,
        })) ?? null,
      candidates: parsedCandidates.map((candidate, idx) => ({
        source: candidate.source,
        index: candidate.index,
        selectedByTokenizer: selectedIndexSet.has(idx),
        tokenCount: candidate.tokens.length,
        tokens: candidate.tokens,
      })),
    };

    if (args.emitJson) {
      process.stdout.write(`${JSON.stringify(payload, null, args.emitPretty ? 2 : undefined)}\n`);
    } else {
      renderTextOutput(payload);
    }
  } finally {
    await shutdownYomitanRuntime(yomitan);
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
