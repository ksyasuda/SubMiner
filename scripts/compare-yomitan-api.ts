import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';

import { createTokenizerDepsRuntime, tokenizeSubtitle } from '../src/core/services/tokenizer.js';
import { selectYomitanParseTokens } from '../src/core/services/tokenizer/parser-selection-stage.js';
import { enrichTokensWithMecabPos1 } from '../src/core/services/tokenizer/parser-enrichment-stage.js';
import { resolveYomitanExtensionPath as resolveBuiltYomitanExtensionPath } from '../src/core/services/yomitan-extension-paths.js';
import { MecabTokenizer } from '../src/mecab-tokenizer.js';
import type { MergedToken } from '../src/types.js';

// Compares SubMiner's full tokenization pipeline against a stock Yomitan
// instance reached through the yomitan-api native-messaging bridge
// (https://github.com/yomidevs/yomitan-api, default http://127.0.0.1:19633).
// The API response is raw parseText output; it is mapped through SubMiner's
// own selectYomitanParseTokens so both sides share identical mapping
// semantics and the diff isolates real segmentation/headword divergence.

const DEFAULT_YOMITAN_USER_DATA_PATH = path.join(os.homedir(), '.config', 'SubMiner');
const DEFAULT_API_URL = 'http://127.0.0.1:19633';
const DEFAULT_SCAN_LENGTH = 40;
const FIXTURE_DIR = path.join('src', 'core', 'services', 'tokenizer', '__fixtures__', 'golden');

interface CliOptions {
  sentences: string[];
  fromFixtures: boolean;
  apiUrl: string;
  scanLength: number;
  emitJson: boolean;
  nameMatch: boolean;
  yomitanExtensionPath?: string;
  yomitanUserDataPath: string;
  mecabCommand?: string;
  mecabDictionaryPath?: string;
}

function printUsage(): void {
  process.stdout.write(`Usage:
  bun run compare-yomitan-api:electron -- [options] [sentence ...]

  Compares SubMiner tokenization against a stock Yomitan instance via the
  yomitan-api bridge. Without sentences, compares every golden-corpus fixture
  text plus any --file lines.

  --from-fixtures            Include golden-corpus fixture texts (default when
                             no sentences are given).
  --file <path>              Read additional sentences, one per line.
  --api-url <url>            yomitan-api base URL (default: ${DEFAULT_API_URL}).
  --scan-length <n>          Scan length for the API request (default: ${DEFAULT_SCAN_LENGTH}).
  --name-match               Keep SubMiner's character-name pre-pass enabled.
                             Off by default: stock Yomitan has no name regions,
                             so it would only add noise to the diff.
  --json                     Emit machine-readable JSON instead of a report.
  --yomitan-extension <path> Path to built Yomitan extension directory.
  --yomitan-user-data <path> Electron userData directory (default: ~/.config/SubMiner).
  --mecab-command <path>     MeCab binary (default: mecab).
  --mecab-dictionary <path>  MeCab dictionary directory.
  -h, --help                 Show usage.

  Exits 1 when any sentence diverges, 0 when everything matches.
`);
}

function requireValue(args: string[], flag: string): string {
  const next = args.shift();
  if (!next) {
    throw new Error(`Missing value for ${flag}`);
  }
  return next;
}

function parseCliArgs(argv: string[]): CliOptions {
  const args = [...argv];
  const sentences: string[] = [];
  let fromFixtures = false;
  let apiUrl = DEFAULT_API_URL;
  let scanLength = DEFAULT_SCAN_LENGTH;
  let emitJson = false;
  let nameMatch = false;
  let yomitanExtensionPath: string | undefined;
  let yomitanUserDataPath = DEFAULT_YOMITAN_USER_DATA_PATH;
  let mecabCommand: string | undefined;
  let mecabDictionaryPath: string | undefined;

  while (args.length > 0) {
    const arg = args.shift();
    if (!arg) break;

    if (arg === '--help' || arg === '-h') {
      printUsage();
      process.exit(0);
    }
    if (arg === '--from-fixtures') {
      fromFixtures = true;
      continue;
    }
    if (arg === '--file') {
      const filePath = requireValue(args, arg);
      const lines = fs
        .readFileSync(filePath, 'utf8')
        .split('\n')
        .map((line) => line.trim())
        .filter((line) => line.length > 0);
      sentences.push(...lines);
      continue;
    }
    if (arg === '--api-url') {
      apiUrl = requireValue(args, arg).replace(/\/$/, '');
      continue;
    }
    if (arg === '--scan-length') {
      const value = Number(requireValue(args, arg));
      if (!Number.isInteger(value) || value <= 0) {
        throw new Error('Invalid --scan-length value');
      }
      scanLength = value;
      continue;
    }
    if (arg === '--json') {
      emitJson = true;
      continue;
    }
    if (arg === '--name-match') {
      nameMatch = true;
      continue;
    }
    if (arg === '--yomitan-extension') {
      yomitanExtensionPath = requireValue(args, arg);
      continue;
    }
    if (arg === '--yomitan-user-data') {
      yomitanUserDataPath = requireValue(args, arg);
      continue;
    }
    if (arg === '--mecab-command') {
      mecabCommand = requireValue(args, arg);
      continue;
    }
    if (arg === '--mecab-dictionary') {
      mecabDictionaryPath = requireValue(args, arg);
      continue;
    }
    if (arg.startsWith('-') && arg !== '-') {
      throw new Error(`Unknown flag: ${arg}`);
    }
    sentences.push(arg);
  }

  if (sentences.length === 0) {
    fromFixtures = true;
  }

  return {
    sentences,
    fromFixtures,
    apiUrl,
    scanLength,
    emitJson,
    nameMatch,
    yomitanExtensionPath,
    yomitanUserDataPath,
    mecabCommand,
    mecabDictionaryPath,
  };
}

function loadFixtureSentences(): string[] {
  if (!fs.existsSync(FIXTURE_DIR)) {
    return [];
  }
  return fs
    .readdirSync(FIXTURE_DIR)
    .filter((entry) => entry.endsWith('.json'))
    .sort()
    .map((entry) => {
      const fixture = JSON.parse(fs.readFileSync(path.join(FIXTURE_DIR, entry), 'utf8')) as {
        input?: { text?: string };
      };
      return fixture.input?.text ?? '';
    })
    .filter((text) => text.length > 0);
}

function normalizeTokenizerText(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\\N/g, '\n')
    .replace(/\\n/g, '\n')
    .trim()
    .replace(/\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

interface FuriganaChunk {
  startPos: number;
  endPos: number;
  reading: string;
  headword: string;
}

// Walk the raw parseText content the same way the parse-result mapper does
// (whole segments, char offsets accumulated across lines) and capture, per
// chunk, the effective reading (kana chunks read as themselves) and the first
// dictionary headword term. The mapper's token spans are authoritative for
// grouping; these chunks supply the complete reading/lemma the legacy mapper
// does not synthesize.
function collectFuriganaChunks(parseResults: unknown): FuriganaChunk[] {
  const chunks: FuriganaChunk[] = [];
  if (!Array.isArray(parseResults)) {
    return chunks;
  }
  const content = (parseResults[0] as { content?: unknown } | undefined)?.content;
  if (!Array.isArray(content)) {
    return chunks;
  }
  let charOffset = 0;
  for (const line of content) {
    if (!Array.isArray(line)) {
      continue;
    }
    for (const segment of line as Array<Record<string, unknown>>) {
      const text = typeof segment?.text === 'string' ? segment.text : '';
      if (!text) {
        continue;
      }
      const reading =
        typeof segment.reading === 'string' && segment.reading.length > 0 ? segment.reading : text;
      let headword = '';
      const headwords = segment.headwords;
      if (Array.isArray(headwords)) {
        for (const row of headwords) {
          const first = Array.isArray(row) ? (row[0] as { term?: unknown } | undefined) : undefined;
          if (first && typeof first.term === 'string' && first.term.trim().length > 0) {
            headword = first.term.trim();
            break;
          }
        }
      }
      chunks.push({ startPos: charOffset, endPos: charOffset + text.length, reading, headword });
      charOffset += text.length;
    }
  }
  return chunks;
}

function overlayChunkMetadata(tokens: MergedToken[], chunks: FuriganaChunk[]): MergedToken[] {
  return tokens.map((token) => {
    const covered = chunks.filter(
      (chunk) => chunk.startPos >= token.startPos && chunk.endPos <= token.endPos,
    );
    if (covered.length === 0) {
      return token;
    }
    const reading = covered.map((chunk) => chunk.reading).join('');
    const headword = covered.find((chunk) => chunk.headword)?.headword ?? token.surface;
    return { ...token, reading, headword };
  });
}

async function fetchApiTokens(
  apiUrl: string,
  text: string,
  scanLength: number,
): Promise<MergedToken[] | null> {
  const response = await fetch(`${apiUrl}/tokenize`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ text, scanLength, parser: 'scanning-parser' }),
  });
  if (!response.ok) {
    throw new Error(`yomitan-api /tokenize responded ${response.status}`);
  }
  const parseResults = (await response.json()) as unknown;
  const tokens = selectYomitanParseTokens(parseResults, () => false, 'headword');
  if (!tokens) {
    return tokens;
  }
  return overlayChunkMetadata(tokens, collectFuriganaChunks(parseResults));
}

interface ComparedToken {
  surface: string;
  reading: string;
  headword: string;
  startPos: number;
  endPos: number;
  frequencyRank?: number;
}

function toComparedTokens(tokens: MergedToken[] | null): ComparedToken[] {
  return (tokens ?? []).map((token) => ({
    surface: token.surface,
    reading: token.reading,
    headword: token.headword,
    startPos: token.startPos,
    endPos: token.endPos,
    frequencyRank: token.frequencyRank,
  }));
}

interface MismatchRegion {
  subminer: ComparedToken[];
  yomitan: ComparedToken[];
}

interface PairDiff {
  surface: string;
  field: 'headword' | 'reading';
  subminer: string;
  yomitan: string;
}

interface SentenceComparison {
  text: string;
  subminerTokens: ComparedToken[];
  yomitanTokens: ComparedToken[];
  segmentationMatches: boolean;
  mismatchRegions: MismatchRegion[];
  pairDiffs: PairDiff[];
  error?: string;
}

// Walk both token lists in span order. Tokens whose [startPos, endPos) agree
// are paired; disagreeing runs are grouped until the boundaries realign.
function compareTokenLists(
  subminer: ComparedToken[],
  yomitan: ComparedToken[],
): Pick<SentenceComparison, 'segmentationMatches' | 'mismatchRegions' | 'pairDiffs'> {
  const mismatchRegions: MismatchRegion[] = [];
  const pairDiffs: PairDiff[] = [];
  let i = 0;
  let j = 0;

  while (i < subminer.length || j < yomitan.length) {
    const a = subminer[i];
    const b = yomitan[j];
    if (a && b && a.startPos === b.startPos && a.endPos === b.endPos) {
      if (a.headword !== b.headword) {
        pairDiffs.push({
          surface: a.surface,
          field: 'headword',
          subminer: a.headword,
          yomitan: b.headword,
        });
      }
      // An empty reading means "reads as its surface" (kana tokens, unparsed
      // runs) — normalize before comparing.
      if ((a.reading || a.surface) !== (b.reading || b.surface)) {
        pairDiffs.push({
          surface: a.surface,
          field: 'reading',
          subminer: a.reading,
          yomitan: b.reading,
        });
      }
      i += 1;
      j += 1;
      continue;
    }

    const region: MismatchRegion = { subminer: [], yomitan: [] };
    let aEnd = a ? a.startPos : Number.MAX_SAFE_INTEGER;
    let bEnd = b ? b.startPos : Number.MAX_SAFE_INTEGER;
    do {
      if (aEnd <= bEnd && i < subminer.length) {
        const token = subminer[i];
        if (token) {
          region.subminer.push(token);
          aEnd = token.endPos;
        }
        i += 1;
      } else if (j < yomitan.length) {
        const token = yomitan[j];
        if (token) {
          region.yomitan.push(token);
          bEnd = token.endPos;
        }
        j += 1;
      } else {
        break;
      }
    } while (aEnd !== bEnd && (i < subminer.length || j < yomitan.length));
    mismatchRegions.push(region);
  }

  return {
    segmentationMatches: mismatchRegions.length === 0,
    mismatchRegions,
    pairDiffs,
  };
}

function renderTokenRun(tokens: ComparedToken[]): string {
  if (tokens.length === 0) {
    return '(no tokens)';
  }
  return tokens.map((token) => token.surface).join('|');
}

function renderReport(comparisons: SentenceComparison[]): void {
  let divergent = 0;
  for (const comparison of comparisons) {
    const clean =
      !comparison.error && comparison.segmentationMatches && comparison.pairDiffs.length === 0;
    if (!clean) {
      divergent += 1;
    }
    process.stdout.write(`\n${clean ? 'MATCH' : 'DIFF '} ${comparison.text}\n`);
    if (comparison.error) {
      process.stdout.write(`  error: ${comparison.error}\n`);
      continue;
    }
    if (clean) {
      process.stdout.write(
        `  ${comparison.subminerTokens.length} tokens: ${renderTokenRun(comparison.subminerTokens)}\n`,
      );
      continue;
    }
    if (!comparison.segmentationMatches) {
      process.stdout.write('  segmentation:\n');
      for (const region of comparison.mismatchRegions) {
        process.stdout.write(`    subminer: ${renderTokenRun(region.subminer)}\n`);
        process.stdout.write(`    yomitan:  ${renderTokenRun(region.yomitan)}\n`);
      }
    }
    for (const diff of comparison.pairDiffs) {
      process.stdout.write(
        `  ${diff.surface}: ${diff.field} "${diff.subminer}" (subminer) vs "${diff.yomitan}" (yomitan)\n`,
      );
    }
  }
  process.stdout.write(
    `\n${comparisons.length - divergent}/${comparisons.length} sentences match stock Yomitan\n`,
  );
}

async function main(): Promise<void> {
  const options = parseCliArgs(process.argv.slice(2));
  const sentences = [...(options.fromFixtures ? loadFixtureSentences() : []), ...options.sentences];
  if (sentences.length === 0) {
    throw new Error('No sentences to compare (no fixtures found and none provided).');
  }

  // Fail fast with a clear message when the bridge is not running.
  try {
    const probe = await fetch(`${options.apiUrl}/serverVersion`, {
      method: 'POST',
      body: '{}',
    });
    if (!probe.ok) {
      throw new Error(`status ${probe.status}`);
    }
  } catch (error) {
    throw new Error(
      `yomitan-api is not reachable at ${options.apiUrl} (${(error as Error).message}). ` +
        'Ensure the browser is running and "Enable Yomitan API" is on in Yomitan settings.',
    );
  }

  const electronModule = await import('electron').catch(() => null);
  if (!electronModule?.app || !electronModule?.session) {
    throw new Error('Electron runtime required; run via `bun run compare-yomitan-api:electron`.');
  }

  const mecabTokenizer = new MecabTokenizer({
    mecabCommand: options.mecabCommand,
    dictionaryPath: options.mecabDictionaryPath,
  });
  if (!(await mecabTokenizer.checkAvailability())) {
    throw new Error('MeCab is not available on this system.');
  }

  if (typeof electronModule.app.setPath === 'function') {
    electronModule.app.setPath('userData', options.yomitanUserDataPath);
  }
  await electronModule.app.whenReady();

  const extensionPath = resolveBuiltYomitanExtensionPath({
    explicitPath: options.yomitanExtensionPath,
    cwd: process.cwd(),
  });
  if (!extensionPath) {
    throw new Error('No built Yomitan extension found; run `bun run build:yomitan` first.');
  }
  const extension = await electronModule.session.defaultSession.loadExtension(extensionPath, {
    allowFileAccess: true,
  });

  let parserWindow: Electron.BrowserWindow | null = null;
  let parserReadyPromise: Promise<void> | null = null;
  let parserInitPromise: Promise<boolean> | null = null;

  const runtimeDeps = createTokenizerDepsRuntime({
    getYomitanExt: () => extension,
    getYomitanParserWindow: () => parserWindow,
    setYomitanParserWindow: (window) => {
      parserWindow = window;
    },
    getYomitanParserReadyPromise: () => parserReadyPromise,
    setYomitanParserReadyPromise: (promise) => {
      parserReadyPromise = promise;
    },
    getYomitanParserInitPromise: () => parserInitPromise,
    setYomitanParserInitPromise: (promise) => {
      parserInitPromise = promise;
    },
    isKnownWord: () => false,
    getKnownWordMatchMode: () => 'headword',
    getJlptLevel: () => null,
    getNameMatchEnabled: () => options.nameMatch,
    getMecabTokenizer: () => ({
      tokenize: (text: string) => mecabTokenizer.tokenize(text),
    }),
  });
  const deps = {
    ...runtimeDeps,
    enrichTokensWithMecab: async (
      tokens: Parameters<typeof enrichTokensWithMecabPos1>[0],
      mecabTokens: Parameters<typeof enrichTokensWithMecabPos1>[1],
    ) => enrichTokensWithMecabPos1(tokens, mecabTokens),
  };

  const comparisons: SentenceComparison[] = [];
  try {
    for (const sentence of sentences) {
      const text = normalizeTokenizerText(sentence);
      try {
        const [subtitleData, apiTokens] = await Promise.all([
          tokenizeSubtitle(text, deps),
          fetchApiTokens(options.apiUrl, text, options.scanLength),
        ]);
        const subminerTokens = toComparedTokens(subtitleData.tokens);
        const yomitanTokens = toComparedTokens(apiTokens);
        comparisons.push({
          text,
          subminerTokens,
          yomitanTokens,
          ...compareTokenLists(subminerTokens, yomitanTokens),
        });
      } catch (error) {
        comparisons.push({
          text,
          subminerTokens: [],
          yomitanTokens: [],
          segmentationMatches: false,
          mismatchRegions: [],
          pairDiffs: [],
          error: (error as Error).message,
        });
      }
    }
  } finally {
    if (parserWindow) {
      const window = parserWindow as Electron.BrowserWindow;
      if (!window.isDestroyed()) {
        window.destroy();
      }
    }
    electronModule.app.quit();
  }

  if (options.emitJson) {
    process.stdout.write(`${JSON.stringify(comparisons, null, 2)}\n`);
  } else {
    renderReport(comparisons);
  }

  const anyDivergence = comparisons.some(
    (comparison) =>
      comparison.error !== undefined ||
      !comparison.segmentationMatches ||
      comparison.pairDiffs.length > 0,
  );
  process.exitCode = anyDivergence ? 1 : 0;
}

main()
  .then(() => {
    process.exit(process.exitCode ?? 0);
  })
  .catch((error) => {
    console.error(`Error: ${(error as Error).message}`);
    process.exit(1);
  });
