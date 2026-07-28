import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';

import {
  KnownWordCacheManager,
  getKnownWordCacheLifecycleConfig,
} from '../src/anki-integration/known-word-cache.js';
import { getMatureIntervalThresholdDays } from '../src/anki-integration/known-word-maturity.js';
import { resolveConfigDir } from '../src/config/path-resolution.js';
import { ConfigService } from '../src/config/service.js';
import { parseSubtitleCues } from '../src/core/services/subtitle-cue-parser.js';
import { createTokenizerDepsRuntime, tokenizeSubtitle } from '../src/core/services/tokenizer.js';
import {
  resolveCompleteTokenReading,
  resolveKnownWordReadingForMatch,
  resolveKnownWordText,
} from '../src/core/services/tokenizer/annotation-stage.js';
import { MecabTokenizer } from '../src/mecab-tokenizer.js';
import type { MergedToken } from '../src/types.js';
import type { KnownWordMaturityTier } from '../src/types/subtitle.js';
import {
  createYomitanRuntimeStateWithSearch,
  destroyParserWindow,
  loadElectronModule,
  withTimeout,
  type YomitanRuntimeState,
} from './yomitan-script-runtime.js';

interface CliOptions {
  input: string;
  configDir?: string;
  yomitanUserDataPath?: string;
  yomitanExtensionPath?: string;
  limit: number;
  audit: boolean;
  refresh: boolean;
  json: boolean;
  quiet: boolean;
  profileCopy: boolean;
}

type TierOrFallback = KnownWordMaturityTier | 'known-no-tier';

interface TokenReport {
  cueIndex: number;
  startTime: number;
  surface: string;
  headword: string;
  reading: string;
  tier: TierOrFallback;
  noteIds: number[];
}

interface AuditMismatch extends TokenReport {
  liveTier: KnownWordMaturityTier | 'no-notes';
  intervals: number[];
}

const TIERS: readonly KnownWordMaturityTier[] = ['new', 'learning', 'young', 'mature'];
const FALLBACK_TIER_COLORS: Record<KnownWordMaturityTier, string> = {
  new: '#ee99a0',
  learning: '#b7bdf8',
  young: '#91d7e3',
  mature: '#a6da95',
};

function parseCliArgs(argv: string[]): CliOptions {
  const options: CliOptions = {
    input: '',
    limit: 0,
    audit: false,
    refresh: false,
    json: false,
    quiet: false,
    profileCopy: false,
  };
  const rest: string[] = [];

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]!;
    const takeValue = (flag: string): string => {
      const next = argv[i + 1];
      if (!next) {
        throw new Error(`Missing value for ${flag}`);
      }
      i += 1;
      return next;
    };

    if (arg === '--help' || arg === '-h') {
      process.stdout.write(`${usage()}\n`);
      process.exit(0);
    } else if (arg === '--input') {
      options.input = takeValue(arg);
    } else if (arg === '--config-dir') {
      options.configDir = takeValue(arg);
    } else if (arg === '--yomitan-user-data') {
      options.yomitanUserDataPath = takeValue(arg);
    } else if (arg === '--yomitan-extension-path') {
      options.yomitanExtensionPath = takeValue(arg);
    } else if (arg === '--limit') {
      options.limit = Math.max(0, Number.parseInt(takeValue(arg), 10) || 0);
    } else if (arg === '--audit') {
      options.audit = true;
    } else if (arg === '--refresh') {
      options.refresh = true;
    } else if (arg === '--json') {
      options.json = true;
    } else if (arg === '--quiet') {
      options.quiet = true;
    } else if (arg === '--profile-copy') {
      options.profileCopy = true;
    } else if (arg === '--') {
      // `bun run <script> -- --flag ...` forwards the separator too.
      continue;
    } else if (arg.startsWith('--')) {
      throw new Error(`Unknown flag: ${arg}`);
    } else {
      rest.push(arg);
    }
  }

  if (!options.input && rest.length > 0) {
    options.input = rest.join(' ');
  }
  if (!options.input) {
    throw new Error(`No subtitle file given.\n${usage()}`);
  }
  return options;
}

function usage(): string {
  return [
    'Usage: verify-known-word-highlights <subtitle.srt|.ass> [flags]',
    '',
    '  --limit <n>                 Only check the first n cues (default: all)',
    '  --audit                     Re-derive every highlighted tier from live Anki card data',
    '  --refresh                   Force a known-word cache refresh before checking',
    '  --json                      Emit a machine-readable report',
    '  --quiet                     Skip the per-line colored dump',
    '  --profile-copy              Copy the Yomitan profile to a scratch dir so this can run',
    '                              while SubMiner is open (Electron locks the userData dir)',
    '  --config-dir <dir>          SubMiner config dir (default: auto-detected)',
    '  --yomitan-user-data <dir>   Electron userData dir holding the Yomitan profile',
    '  --yomitan-extension-path <dir>',
  ].join('\n');
}

const ANSI_RESET = '\u001b[0m';

function colorize(text: string, hex: string): string {
  const normalized = hex.trim().replace(/^#/, '');
  const expanded =
    normalized.length === 3
      ? normalized
          .split('')
          .map((char) => `${char}${char}`)
          .join('')
      : normalized;
  if (!/^[0-9a-fA-F]{6}$/.test(expanded)) {
    return text;
  }
  const r = Number.parseInt(expanded.slice(0, 2), 16);
  const g = Number.parseInt(expanded.slice(2, 4), 16);
  const b = Number.parseInt(expanded.slice(4, 6), 16);
  return `\u001b[38;2;${r};${g};${b}m${text}${ANSI_RESET}`;
}

function formatTimestamp(seconds: number): string {
  const total = Math.max(0, Math.floor(seconds));
  const mm = String(Math.floor(total / 60)).padStart(2, '0');
  const ss = String(total % 60).padStart(2, '0');
  return `${mm}:${ss}`;
}

// The user's real config dir is copied into a scratch dir so this read-only
// check can never rewrite config.jsonc (ConfigService migrates on load) or the
// live known-word cache (a --refresh persists tier data).
function createScratchState(configDir: string): { dir: string; cachePath: string } {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-highlight-verify-'));
  for (const fileName of ['config.jsonc', 'config.json']) {
    const source = path.join(configDir, fileName);
    if (fs.existsSync(source)) {
      fs.copyFileSync(source, path.join(dir, fileName));
    }
  }
  const cachePath = path.join(dir, 'known-words-cache.json');
  const liveCachePath = path.join(configDir, 'known-words-cache.json');
  if (fs.existsSync(liveCachePath)) {
    fs.copyFileSync(liveCachePath, cachePath);
  }
  return { dir, cachePath };
}

// Electron locks a userData dir, so the Yomitan profile can't be shared with a
// running SubMiner. Copying the dictionary-bearing parts of the profile lets
// this check run mid-session (IndexedDB alone is often over 1 GB).
const YOMITAN_PROFILE_DIRS = [
  'extensions',
  'IndexedDB',
  'Local Extension Settings',
  'Local Storage',
];

function copyYomitanProfile(sourceUserDataPath: string): string {
  const target = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-highlight-profile-'));
  for (const name of YOMITAN_PROFILE_DIRS) {
    const source = path.join(sourceUserDataPath, name);
    if (fs.existsSync(source)) {
      fs.cpSync(source, path.join(target, name), { recursive: true });
    }
  }
  return target;
}

function readPersistedCacheScope(cachePath: string): string | null {
  try {
    const parsed = JSON.parse(fs.readFileSync(cachePath, 'utf-8')) as { scope?: unknown };
    return typeof parsed.scope === 'string' ? parsed.scope : null;
  } catch {
    return null;
  }
}

const ANKI_REQUEST_TIMEOUT_MS = 30_000;

function createAnkiClient(url: string) {
  const request = async (action: string, params: unknown): Promise<unknown> => {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action, version: 6, params }),
      signal: AbortSignal.timeout(ANKI_REQUEST_TIMEOUT_MS),
    });
    if (!response.ok) {
      throw new Error(`AnkiConnect ${action}: HTTP ${response.status} ${response.statusText}`);
    }
    const payload = (await response.json()) as { result: unknown; error: string | null };
    if (payload.error) {
      throw new Error(`AnkiConnect ${action}: ${payload.error}`);
    }
    return payload.result;
  };

  return {
    request,
    findNotes: (query: string) => request('findNotes', { query }),
    notesInfo: (noteIds: number[]) => request('notesInfo', { notes: noteIds }),
  };
}

function resolveTokenMatch(
  token: MergedToken,
  cache: KnownWordCacheManager,
  matchMode: 'surface' | 'headword',
): { tier: KnownWordMaturityTier | null; noteIds: Set<number> } {
  const matchText = resolveKnownWordText(token.surface, token.headword, matchMode);
  const matchReading = resolveKnownWordReadingForMatch(token, matchMode);
  const primaryTier = matchText ? cache.getKnownWordTier(matchText, matchReading) : null;
  if (primaryTier) {
    return { tier: primaryTier, noteIds: cache.getKnownWordMatchNoteIds(matchText, matchReading) };
  }

  const fallbackReading = resolveCompleteTokenReading(token);
  if (!fallbackReading || fallbackReading === matchText.trim()) {
    return {
      tier: null,
      noteIds: matchText ? cache.getKnownWordMatchNoteIds(matchText, matchReading) : new Set(),
    };
  }
  const fallbackOptions = { allowReadingOnlyMatch: false } as const;
  return {
    tier: cache.getKnownWordTier(fallbackReading, undefined, fallbackOptions),
    noteIds: cache.getKnownWordMatchNoteIds(fallbackReading, undefined, fallbackOptions),
  };
}

// Ground truth straight from card data, independent of the Anki search filters
// the cache refresh uses (prop:ivl / is:learn).
function classifyCardsIntoTier(
  cards: Array<{ interval: number; queue: number; type: number }>,
  thresholdDays: number,
): KnownWordMaturityTier {
  if (cards.some((card) => card.interval >= thresholdDays)) return 'mature';
  if (cards.some((card) => card.interval >= 1)) return 'young';
  if (
    cards.some((card) => card.type === 1 || card.type === 3 || card.queue === 1 || card.queue === 4)
  )
    return 'learning';
  return 'new';
}

async function auditTokens(
  reports: TokenReport[],
  client: ReturnType<typeof createAnkiClient>,
  thresholdDays: number,
): Promise<{ mismatches: AuditMismatch[]; auditedNotes: number }> {
  const noteIds = [...new Set(reports.flatMap((report) => report.noteIds))];
  const cardIdsByNote = new Map<number, number[]>();
  for (let i = 0; i < noteIds.length; i += 500) {
    const infos = (await client.notesInfo(noteIds.slice(i, i + 500))) as Array<{
      noteId: number;
      cards?: number[];
    }>;
    for (const info of infos) {
      cardIdsByNote.set(info.noteId, info.cards ?? []);
    }
  }

  const allCardIds = [...cardIdsByNote.values()].flat();
  const cardById = new Map<number, { interval: number; queue: number; type: number }>();
  for (let i = 0; i < allCardIds.length; i += 500) {
    const infos = (await client.request('cardsInfo', {
      cards: allCardIds.slice(i, i + 500),
    })) as Array<{ cardId: number; interval: number; queue: number; type: number }>;
    for (const info of infos) {
      cardById.set(info.cardId, {
        interval: info.interval,
        queue: info.queue,
        type: info.type,
      });
    }
  }

  const mismatches: AuditMismatch[] = [];
  for (const report of reports) {
    const cards = report.noteIds
      .flatMap((noteId) => cardIdsByNote.get(noteId) ?? [])
      .map((cardId) => cardById.get(cardId))
      .filter((card): card is { interval: number; queue: number; type: number } => Boolean(card));
    const liveTier = cards.length === 0 ? 'no-notes' : classifyCardsIntoTier(cards, thresholdDays);
    if (liveTier !== report.tier) {
      mismatches.push({
        ...report,
        liveTier,
        intervals: cards.map((card) => card.interval),
      });
    }
  }
  return { mismatches, auditedNotes: noteIds.length };
}

async function main(): Promise<void> {
  const args = parseCliArgs(process.argv.slice(2));
  let electronModule: typeof import('electron') | null = null;
  let yomitanState: YomitanRuntimeState | null = null;
  let scratchDir: string | null = null;
  let profileCopyDir: string | null = null;

  try {
    const configDir =
      args.configDir ??
      resolveConfigDir({
        homeDir: os.homedir(),
        xdgConfigHome: process.env.XDG_CONFIG_HOME,
        existsSync: fs.existsSync,
      });
    const scratch = createScratchState(configDir);
    scratchDir = scratch.dir;
    const config = new ConfigService(scratch.dir).getConfig();
    const ankiConfig = config.ankiConnect;
    const matchMode = ankiConfig.knownWords?.matchMode === 'surface' ? 'surface' : 'headword';
    const thresholdDays = getMatureIntervalThresholdDays(ankiConfig);
    const client = createAnkiClient(ankiConfig.url);
    const cacheScopeKey = getKnownWordCacheLifecycleConfig(ankiConfig);

    const cache = new KnownWordCacheManager({
      client: { findNotes: (query) => client.findNotes(query), notesInfo: client.notesInfo },
      getConfig: () => ankiConfig,
      knownWordCacheStatePath: scratch.cachePath,
      showStatusNotification: () => {},
    });
    // A cache whose persisted scope key no longer matches the config is
    // discarded on load, so every token would come back unknown.
    if (!args.refresh && readPersistedCacheScope(scratch.cachePath) !== cacheScopeKey) {
      process.stderr.write(
        'warning: the persisted known-word cache was built under different settings and will be ' +
          'ignored (the app refetches it on its next refresh). Re-run with --refresh to fetch tiers now.\n',
      );
    }
    // startLifecycle loads the persisted cache; the refresh timer it arms is
    // cleared before the event loop can run it.
    cache.startLifecycle();
    cache.stopLifecycle();
    if (args.refresh) {
      await cache.refresh(true);
    }

    const cues = parseSubtitleCues(fs.readFileSync(args.input, 'utf-8'), args.input);
    const selectedCues = args.limit > 0 ? cues.slice(0, args.limit) : cues;

    const mecabTokenizer = new MecabTokenizer();
    if (!(await mecabTokenizer.checkAvailability())) {
      throw new Error('MeCab is not available; tokenization would not match the overlay.');
    }
    electronModule = await loadElectronModule();
    const userDataPath = args.profileCopy
      ? copyYomitanProfile(args.yomitanUserDataPath ?? configDir)
      : (args.yomitanUserDataPath ?? configDir);
    profileCopyDir = args.profileCopy ? userDataPath : null;
    if (electronModule?.app && typeof electronModule.app.setPath === 'function') {
      electronModule.app.setPath('userData', userDataPath);
    }
    yomitanState = await createYomitanRuntimeStateWithSearch(
      userDataPath,
      args.yomitanExtensionPath,
    );
    if (!yomitanState.available) {
      throw new Error(`Yomitan tokenizer unavailable: ${yomitanState.note ?? 'unknown reason'}`);
    }

    const deps = createTokenizerDepsRuntime({
      getYomitanExt: () => yomitanState!.yomitanExt as never,
      getYomitanSession: () => yomitanState!.yomitanSession as never,
      getYomitanParserWindow: () => yomitanState!.parserWindow as never,
      setYomitanParserWindow: (window) => {
        yomitanState!.parserWindow = window;
      },
      getYomitanParserReadyPromise: () => yomitanState!.parserReadyPromise as never,
      setYomitanParserReadyPromise: (promise) => {
        yomitanState!.parserReadyPromise = promise;
      },
      getYomitanParserInitPromise: () => yomitanState!.parserInitPromise as never,
      setYomitanParserInitPromise: (promise) => {
        yomitanState!.parserInitPromise = promise;
      },
      isKnownWord: (text, reading, options) => cache.isKnownWord(text, reading, options),
      getKnownWordTier: (text, reading, options) => cache.getKnownWordTier(text, reading, options),
      getKnownWordMatchMode: () => matchMode,
      getKnownWordsEnabled: () => true,
      // Other annotation layers are off so every colored token below is a
      // known-word decision, not an N+1/frequency/name override.
      getNPlusOneEnabled: () => false,
      getNameMatchEnabled: () => false,
      getJlptEnabled: () => false,
      getFrequencyDictionaryEnabled: () => false,
      getJlptLevel: () => null,
      getMecabTokenizer: () => ({ tokenize: (text: string) => mecabTokenizer.tokenize(text) }),
    });

    const styleColors = {
      ...FALLBACK_TIER_COLORS,
      ...(config.subtitleStyle?.knownWordMaturityColors ?? {}),
    } as Record<KnownWordMaturityTier, string>;
    const knownWordColor = config.subtitleStyle?.knownWordColor ?? '#a6da95';

    const reports: TokenReport[] = [];
    const tierCounts: Record<string, number> = {};
    let knownTokens = 0;
    let totalTokens = 0;
    const lines: string[] = [];

    for (const [cueIndex, cue] of selectedCues.entries()) {
      const { text, tokens } = await withTimeout(
        tokenizeSubtitle(cue.text, deps),
        20_000,
        `Tokenizer (cue ${cueIndex + 1})`,
      );
      if (!tokens || tokens.length === 0) {
        continue;
      }

      let cursor = 0;
      let rendered = '';
      const ordered = [...tokens].sort((a, b) => (a.startPos ?? 0) - (b.startPos ?? 0));
      for (const token of ordered) {
        totalTokens += 1;
        const start = Math.min(Math.max(0, token.startPos ?? 0), text.length);
        const end = Math.min(Math.max(start, token.endPos ?? start), text.length);
        if (start > cursor) {
          rendered += text.slice(cursor, start);
        }
        const surfaceText = text.slice(start, end);
        cursor = end;

        if (!token.isKnown) {
          rendered += surfaceText;
          continue;
        }
        knownTokens += 1;
        const tier: TierOrFallback = token.knownMaturity ?? 'known-no-tier';
        tierCounts[tier] = (tierCounts[tier] ?? 0) + 1;
        rendered += colorize(
          surfaceText,
          tier === 'known-no-tier' ? knownWordColor : styleColors[tier],
        );
        reports.push({
          cueIndex,
          startTime: cue.startTime,
          surface: token.surface,
          headword: token.headword,
          reading: token.reading,
          tier,
          noteIds: [...resolveTokenMatch(token, cache, matchMode).noteIds],
        });
      }
      rendered += text.slice(cursor);
      lines.push(`${formatTimestamp(cue.startTime)}  ${rendered}`);
    }

    if (totalTokens === 0 && selectedCues.length > 0) {
      throw new Error(
        'Yomitan returned no tokens. SubMiner is probably running and holding the Electron ' +
          'profile lock - quit it, or re-run with --profile-copy.',
      );
    }

    let audit: { mismatches: AuditMismatch[]; auditedNotes: number } | null = null;
    if (args.audit) {
      audit = await auditTokens(reports, client, thresholdDays);
    }

    if (args.json) {
      process.stdout.write(
        `${JSON.stringify(
          {
            input: args.input,
            cues: selectedCues.length,
            totalTokens,
            knownTokens,
            tierCounts,
            matureThresholdDays: thresholdDays,
            matchMode,
            tokens: reports,
            audit,
          },
          null,
          2,
        )}\n`,
      );
      return;
    }

    if (!args.quiet) {
      process.stdout.write(`${lines.join('\n')}\n\n`);
    }
    process.stdout.write(
      [
        `file           : ${args.input}`,
        `cues checked   : ${selectedCues.length} of ${cues.length}`,
        `tokens         : ${totalTokens} (${knownTokens} known, ${totalTokens - knownTokens} unknown)`,
        `match mode     : ${matchMode}   mature threshold: ${thresholdDays}d`,
        '',
        'known-token tiers:',
        ...TIERS.map(
          (tier) =>
            `  ${colorize(tier.padEnd(9), styleColors[tier])} ${String(tierCounts[tier] ?? 0).padStart(5)}` +
            `  ${styleColors[tier]}`,
        ),
        `  ${colorize('no tier'.padEnd(9), knownWordColor)} ${String(tierCounts['known-no-tier'] ?? 0).padStart(5)}  ${knownWordColor} (falls back to knownWordColor)`,
        '',
      ].join('\n'),
    );

    if (audit) {
      process.stdout.write(
        `audit: ${reports.length - audit.mismatches.length}/${reports.length} highlighted tokens agree with live Anki card data ` +
          `(${audit.auditedNotes} notes)\n`,
      );
      for (const mismatch of audit.mismatches.slice(0, 40)) {
        process.stdout.write(
          `  ${formatTimestamp(mismatch.startTime)} ${mismatch.surface} (${mismatch.headword}) ` +
            `shown=${mismatch.tier} live=${mismatch.liveTier} ivl=[${mismatch.intervals.join(', ')}] ` +
            `notes=[${mismatch.noteIds.join(', ')}]\n`,
        );
      }
      if (audit.mismatches.length > 40) {
        process.stdout.write(`  ... ${audit.mismatches.length - 40} more\n`);
      }
    }
  } finally {
    destroyParserWindow(yomitanState?.parserWindow ?? null);
    for (const dir of [scratchDir, profileCopyDir]) {
      if (dir) {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
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
