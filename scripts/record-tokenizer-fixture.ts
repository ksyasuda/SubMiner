import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';

import { createTokenizerDepsRuntime, tokenizeSubtitle } from '../src/core/services/tokenizer.js';
import {
  buildFixtureLookups,
  classifyInjectedScript,
  fixtureConfigGetters,
  hashInjectedScript,
  projectGoldenTokens,
  pruneGoldenMessages,
} from '../src/core/services/tokenizer/golden-corpus-harness.js';
import type {
  GoldenFixture,
  GoldenFixtureConfig,
  GoldenRecordedMessage,
  GoldenRecordedScript,
} from '../src/core/services/tokenizer/golden-corpus-harness.js';
import { enrichTokensWithMecabPos1 } from '../src/core/services/tokenizer/parser-enrichment-stage.js';
import { resolveYomitanExtensionPath as resolveBuiltYomitanExtensionPath } from '../src/core/services/yomitan-extension-paths.js';
import { MecabTokenizer } from '../src/mecab-tokenizer.js';
import type { JlptLevel, Token } from '../src/types.js';

const DEFAULT_YOMITAN_USER_DATA_PATH = path.join(os.homedir(), '.config', 'SubMiner');
const DEFAULT_FIXTURE_DIR = path.join(
  'src',
  'core',
  'services',
  'tokenizer',
  '__fixtures__',
  'golden',
);
const JLPT_LEVELS: ReadonlySet<string> = new Set(['N1', 'N2', 'N3', 'N4', 'N5']);

interface CliOptions {
  name: string;
  text: string;
  description?: string;
  issueRefs: string[];
  config: GoldenFixtureConfig;
  outDir: string;
  force: boolean;
  yomitanExtensionPath?: string;
  yomitanUserDataPath: string;
  mecabCommand?: string;
  mecabDictionaryPath?: string;
}

function printUsage(): void {
  process.stdout.write(`Usage:
  bun run record-tokenizer-fixture:electron -- --name <slug> [options] <subtitle text>

  Records a golden tokenizer fixture: raw Yomitan backend responses, raw MeCab
  tokens, and the annotated tokens the full pipeline produced for the text.
  Review the "expected" block before committing — it becomes the assertion.

  --name <slug>              Fixture name (file becomes <slug>.json). Required.
  --description <text>       What behavior this fixture pins down.
  --issue <ref>              Related issue/PR ref, repeatable (e.g. --issue "#156").
  --known-word <w[:reading]> Mark a word known, repeatable.
  --jlpt <term=level>        JLPT level for a term, repeatable (level N1..N5).
  --local-frequency <t=rank> Local frequency-dictionary rank, repeatable.
  --match-mode <mode>        Known-word match mode: headword|surface.
  --min-sentence-words <n>   Minimum sentence words for N+1.
  --no-known-words           Disable known-word annotation.
  --no-nplusone              Disable N+1 targeting.
  --no-jlpt                  Disable JLPT annotation.
  --no-name-match            Disable character-name matching.
  --no-frequency             Disable frequency annotation.
  --out-dir <path>           Fixture directory (default: ${DEFAULT_FIXTURE_DIR}).
  --force                    Overwrite an existing fixture file.
  --yomitan-extension <path> Path to built Yomitan extension directory.
  --yomitan-user-data <path> Electron userData directory (default: ~/.config/SubMiner).
  --mecab-command <path>     MeCab binary (default: mecab).
  --mecab-dictionary <path>  MeCab dictionary directory.
  -h, --help                 Show usage.
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
  const inputParts: string[] = [];
  const issueRefs: string[] = [];
  const knownWords: GoldenFixtureConfig['knownWords'] = [];
  const jlptLevels: Record<string, JlptLevel> = {};
  const localFrequencyRanks: Record<string, number> = {};
  const config: GoldenFixtureConfig = {};
  let name: string | undefined;
  let description: string | undefined;
  let outDir = DEFAULT_FIXTURE_DIR;
  let force = false;
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
    if (arg === '--name') {
      name = requireValue(args, arg);
      continue;
    }
    if (arg === '--description') {
      description = requireValue(args, arg);
      continue;
    }
    if (arg === '--issue') {
      issueRefs.push(requireValue(args, arg));
      continue;
    }
    if (arg === '--known-word') {
      const value = requireValue(args, arg);
      const separator = value.indexOf(':');
      knownWords.push(
        separator === -1
          ? value
          : { text: value.slice(0, separator), reading: value.slice(separator + 1) },
      );
      continue;
    }
    if (arg === '--jlpt') {
      const value = requireValue(args, arg);
      const separator = value.indexOf('=');
      const level = separator === -1 ? '' : value.slice(separator + 1);
      if (separator === -1 || !JLPT_LEVELS.has(level)) {
        throw new Error(`Invalid --jlpt value "${value}"; expected <term>=N1..N5`);
      }
      jlptLevels[value.slice(0, separator)] = level as JlptLevel;
      continue;
    }
    if (arg === '--local-frequency') {
      const value = requireValue(args, arg);
      const separator = value.indexOf('=');
      const rank = separator === -1 ? Number.NaN : Number(value.slice(separator + 1));
      if (!Number.isInteger(rank) || rank <= 0) {
        throw new Error(`Invalid --local-frequency value "${value}"; expected <term>=<rank>`);
      }
      localFrequencyRanks[value.slice(0, separator)] = rank;
      continue;
    }
    if (arg === '--match-mode') {
      const value = requireValue(args, arg);
      if (value !== 'headword' && value !== 'surface') {
        throw new Error(`Invalid --match-mode "${value}"; expected headword|surface`);
      }
      config.knownWordMatchMode = value;
      config.frequencyMatchMode = value;
      continue;
    }
    if (arg === '--min-sentence-words') {
      const value = Number(requireValue(args, arg));
      if (!Number.isInteger(value) || value < 0) {
        throw new Error('Invalid --min-sentence-words value');
      }
      config.minSentenceWordsForNPlusOne = value;
      continue;
    }
    if (arg === '--no-known-words') {
      config.knownWordsEnabled = false;
      continue;
    }
    if (arg === '--no-nplusone') {
      config.nPlusOneEnabled = false;
      continue;
    }
    if (arg === '--no-jlpt') {
      config.jlptEnabled = false;
      continue;
    }
    if (arg === '--no-name-match') {
      config.nameMatchEnabled = false;
      continue;
    }
    if (arg === '--no-frequency') {
      config.frequencyEnabled = false;
      continue;
    }
    if (arg === '--out-dir') {
      outDir = requireValue(args, arg);
      continue;
    }
    if (arg === '--force') {
      force = true;
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
    inputParts.push(arg);
  }

  if (!name || !/^[a-z0-9][a-z0-9-]*$/.test(name)) {
    throw new Error('A kebab-case --name is required (e.g. --name kureru-lexical-keep)');
  }
  const text = inputParts.join(' ').trim();
  if (!text) {
    throw new Error('Provide the subtitle text to record as positional arguments.');
  }

  if (knownWords.length > 0) config.knownWords = knownWords;
  if (Object.keys(jlptLevels).length > 0) config.jlptLevels = jlptLevels;
  if (Object.keys(localFrequencyRanks).length > 0) {
    config.localFrequencyRanks = localFrequencyRanks;
  }

  return {
    name,
    text,
    description,
    issueRefs,
    config,
    outDir,
    force,
    yomitanExtensionPath,
    yomitanUserDataPath,
    mecabCommand,
    mecabDictionaryPath,
  };
}

const RECORDER_SHIM_SCRIPT = `
(() => {
  const g = globalThis;
  if (g.__subminerGoldenRecorderInstalled) { return true; }
  g.__subminerGoldenRecorderInstalled = true;
  g.__subminerGoldenMessages = [];
  const runtime = chrome.runtime;
  const original = runtime.sendMessage.bind(runtime);
  runtime.sendMessage = (payload, callback) => {
    return original(payload, (response) => {
      try {
        g.__subminerGoldenMessages.push(JSON.parse(JSON.stringify({
          action: payload && payload.action ? payload.action : '',
          params: payload && typeof payload.params !== 'undefined' ? payload.params : null,
          response: typeof response !== 'undefined' ? response : null,
        })));
      } catch {}
      if (callback) { callback(response); }
    });
  };
  return true;
})();
`;

const RECORDER_DRAIN_SCRIPT = `
(() => {
  const g = globalThis;
  const log = g.__subminerGoldenMessages || [];
  g.__subminerGoldenMessages = [];
  return JSON.stringify(log);
})();
`;

interface RecordingSink {
  messages: GoldenRecordedMessage[];
  scripts: GoldenRecordedScript[];
  mecab: Record<string, Token[] | null>;
}

function patchParserWindowForRecording(
  window: Electron.BrowserWindow | null,
  patched: WeakSet<object>,
  sink: RecordingSink,
): void {
  if (!window || window.isDestroyed()) {
    return;
  }
  const webContents = window.webContents as unknown as {
    executeJavaScript: (script: string, userGesture?: boolean) => Promise<unknown>;
  };
  if (patched.has(webContents)) {
    return;
  }
  patched.add(webContents);

  const original = webContents.executeJavaScript.bind(webContents);
  let shimReady: Promise<unknown> | null = null;
  const drain = async (): Promise<void> => {
    try {
      const raw = (await original(RECORDER_DRAIN_SCRIPT, true)) as string;
      const drained = JSON.parse(raw) as GoldenRecordedMessage[];
      sink.messages.push(...drained);
    } catch (error) {
      process.stderr.write(
        `Warning: failed to drain recorded messages: ${(error as Error).message}\n`,
      );
    }
  };

  webContents.executeJavaScript = async (script: string, userGesture?: boolean) => {
    if (!shimReady) {
      shimReady = original(RECORDER_SHIM_SCRIPT, true);
    }
    try {
      await shimReady;
    } catch {
      shimReady = null;
    }
    let result: unknown;
    try {
      result = await original(script, userGesture);
    } finally {
      await drain();
    }
    sink.scripts.push({
      sha256: hashInjectedScript(script),
      marker: classifyInjectedScript(script),
      result: result === undefined ? null : result,
    });
    return result;
  };
}

async function main(): Promise<void> {
  const options = parseCliArgs(process.argv.slice(2));
  const outPath = path.join(options.outDir, `${options.name}.json`);
  if (fs.existsSync(outPath) && !options.force) {
    throw new Error(`${outPath} already exists; pass --force to overwrite.`);
  }

  const electronModule = await import('electron').catch(() => null);
  if (!electronModule?.app || !electronModule?.session) {
    throw new Error(
      'Electron runtime is required; run via `bun run record-tokenizer-fixture:electron`.',
    );
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

  const sink: RecordingSink = { messages: [], scripts: [], mecab: {} };
  const patched = new WeakSet<object>();
  let parserWindow: Electron.BrowserWindow | null = null;
  let parserReadyPromise: Promise<void> | null = null;
  let parserInitPromise: Promise<boolean> | null = null;

  const lookups = buildFixtureLookups(options.config);
  const runtimeDeps = createTokenizerDepsRuntime({
    getYomitanExt: () => extension,
    getYomitanParserWindow: () => {
      patchParserWindowForRecording(parserWindow, patched, sink);
      return parserWindow;
    },
    setYomitanParserWindow: (window) => {
      parserWindow = window;
      patchParserWindowForRecording(parserWindow, patched, sink);
    },
    getYomitanParserReadyPromise: () => parserReadyPromise,
    setYomitanParserReadyPromise: (promise) => {
      parserReadyPromise = promise;
    },
    getYomitanParserInitPromise: () => parserInitPromise,
    setYomitanParserInitPromise: (promise) => {
      parserInitPromise = promise;
    },
    isKnownWord: lookups.isKnownWord,
    getJlptLevel: lookups.getJlptLevel,
    getFrequencyRank: lookups.getFrequencyRank,
    ...fixtureConfigGetters(options.config),
    getMecabTokenizer: () => ({
      tokenize: async (text: string) => {
        const rawTokens = await mecabTokenizer.tokenize(text);
        sink.mecab[text] = rawTokens;
        return rawTokens;
      },
    }),
  });
  const deps = {
    ...runtimeDeps,
    // Same pinned enrichment implementation the replay harness uses.
    enrichTokensWithMecab: async (
      tokens: Parameters<typeof enrichTokensWithMecabPos1>[0],
      mecabTokens: Parameters<typeof enrichTokensWithMecabPos1>[1],
    ) => enrichTokensWithMecabPos1(tokens, mecabTokens),
  };

  try {
    const subtitleData = await tokenizeSubtitle(options.text, deps);
    const expectedTokens = projectGoldenTokens(subtitleData.tokens);

    const prunedMessages = pruneGoldenMessages(sink.messages);
    const fixture: GoldenFixture = {
      name: options.name,
      ...(options.description ? { description: options.description } : {}),
      ...(options.issueRefs.length > 0 ? { issueRefs: options.issueRefs } : {}),
      recordedAt: new Date().toISOString(),
      input: { text: options.text },
      config: options.config,
      recording: {
        messages: prunedMessages,
        scripts: sink.scripts,
        mecab: sink.mecab,
      },
      expected: { tokens: expectedTokens },
    };

    fs.mkdirSync(options.outDir, { recursive: true });
    fs.writeFileSync(outPath, `${JSON.stringify(fixture, null, 2)}\n`);

    process.stdout.write(`Recorded ${outPath}\n`);
    process.stdout.write(
      `  messages=${prunedMessages.length} scripts=${sink.scripts.length} mecabTexts=${Object.keys(sink.mecab).length}\n`,
    );
    process.stdout.write('\nExpected tokens (review before committing):\n');
    if (!expectedTokens || expectedTokens.length === 0) {
      process.stdout.write('  (none)\n');
    } else {
      for (const token of expectedTokens) {
        const flags = [
          token.isKnown ? 'known' : null,
          token.isNPlusOneTarget ? 'n+1' : null,
          token.isNameMatch ? 'name' : null,
          token.isUnparsedRun ? 'unparsed' : null,
          token.jlptLevel ?? null,
          token.frequencyRank !== undefined ? `freq=${token.frequencyRank}` : null,
        ]
          .filter(Boolean)
          .join(', ');
        process.stdout.write(
          `  ${token.surface} -> ${token.headword} (${token.reading})${flags ? ` [${flags}]` : ''}\n`,
        );
      }
    }
    process.stdout.write(
      '\nVerify replay with: bun test src/core/services/tokenizer/golden-corpus.test.ts\n',
    );
  } finally {
    if (parserWindow) {
      const window = parserWindow as Electron.BrowserWindow;
      if (!window.isDestroyed()) {
        window.destroy();
      }
    }
    electronModule.app.quit();
  }
}

main()
  .then(() => process.exit(0))
  .catch((error) => {
    console.error(`Error: ${(error as Error).message}`);
    process.exit(1);
  });
