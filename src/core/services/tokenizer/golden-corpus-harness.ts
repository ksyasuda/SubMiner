import { createHash } from 'node:crypto';
import vm from 'node:vm';
import type {
  FrequencyDictionaryMatchMode,
  JlptLevel,
  MergedToken,
  NPlusOneMatchMode,
  PartOfSpeech,
  Token,
} from '../../../types';
import { createTokenizerDepsRuntime } from '../tokenizer';
import type { TokenizerServiceDeps } from '../tokenizer';
import { enrichTokensWithMecabPos1 } from './parser-enrichment-stage';

// Golden-corpus fixtures capture a real subtitle line, the raw Yomitan
// backend responses (chrome.runtime.sendMessage level) and raw MeCab tokens
// observed while tokenizing it live, plus the annotated tokens the full
// pipeline produced. Replay runs the real injected scanning scripts in a vm
// against the recorded responses, so merge/enrichment/annotation/filtering
// are exercised end-to-end without Electron or dictionaries.

export interface GoldenFixtureKnownWord {
  text: string;
  reading?: string;
}

export interface GoldenFixtureConfig {
  knownWords?: Array<string | GoldenFixtureKnownWord>;
  knownWordMatchMode?: NPlusOneMatchMode;
  jlptLevels?: Record<string, JlptLevel>;
  localFrequencyRanks?: Record<string, number>;
  knownWordsEnabled?: boolean;
  nPlusOneEnabled?: boolean;
  jlptEnabled?: boolean;
  nameMatchEnabled?: boolean;
  frequencyEnabled?: boolean;
  frequencyMatchMode?: FrequencyDictionaryMatchMode;
  minSentenceWordsForNPlusOne?: number;
}

export interface GoldenRecordedMessage {
  action: string;
  params: unknown;
  /** Raw response object the page callback received: { result } or { error }. */
  response: unknown;
}

export interface GoldenRecordedScript {
  sha256: string;
  /** Short classification of the script, for debugging only. */
  marker: string;
  result: unknown;
}

export interface GoldenFixtureRecording {
  messages: GoldenRecordedMessage[];
  scripts: GoldenRecordedScript[];
  /** Raw MeCab tokens keyed by the exact text passed to tokenizeWithMecab. */
  mecab: Record<string, Token[] | null>;
}

export interface GoldenExpectedToken {
  surface: string;
  reading: string;
  headword: string;
  headwordReading?: string;
  startPos: number;
  endPos: number;
  partOfSpeech: PartOfSpeech;
  pos1?: string;
  pos2?: string;
  pos3?: string;
  isKnown: boolean;
  isNPlusOneTarget: boolean;
  isNameMatch?: true;
  isUnparsedRun?: true;
  jlptLevel?: JlptLevel;
  frequencyRank?: number;
}

export interface GoldenFixture {
  name: string;
  description?: string;
  issueRefs?: string[];
  recordedAt?: string;
  input: { text: string };
  config: GoldenFixtureConfig;
  recording: GoldenFixtureRecording;
  expected: { tokens: GoldenExpectedToken[] | null };
}

// Only these backend actions are requested by the scanning/frequency scripts;
// everything else in the recorded page traffic is Yomitan's own chatter.
const GOLDEN_MESSAGE_ACTIONS: ReadonlySet<string> = new Set([
  'optionsGetFull',
  'getDictionaryInfo',
  'parseText',
  'termsFind',
  'getTermFrequencies',
]);

// Field candidates appendDictionaryNames() in the injected helpers reads off
// definitions/pronunciations — the only reason those arrays are consulted.
const DICTIONARY_NAME_FIELDS = [
  'dictionary',
  'dictionaryName',
  'name',
  'title',
  'dictionaryTitle',
  'dictionaryAlias',
] as const;

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function pruneToDictionaryNameFields(value: unknown): unknown {
  if (!isRecord(value)) {
    return {};
  }
  const pruned: Record<string, unknown> = {};
  for (const field of DICTIONARY_NAME_FIELDS) {
    if (typeof value[field] === 'string') {
      pruned[field] = value[field];
    }
  }
  return pruned;
}

// Fields getBestFrequencyRank() and the frequency grouping helpers read off a
// termsFind frequency item, beyond the dictionary-name candidates.
const FREQUENCY_ITEM_FIELDS = [
  'headwordIndex',
  'displayValue',
  'displayValueParsed',
  'frequency',
  'dictionaryIndex',
  'term',
  'reading',
] as const;

function pruneFrequencyItem(item: unknown): unknown {
  if (!isRecord(item)) {
    return item;
  }
  const pruned = pruneToDictionaryNameFields(item) as Record<string, unknown>;
  for (const field of FREQUENCY_ITEM_FIELDS) {
    if (item[field] !== undefined) {
      pruned[field] = item[field];
    }
  }
  return pruned;
}

function pruneTermsFindEntry(entry: unknown): unknown {
  if (!isRecord(entry)) {
    return entry;
  }
  const pruned: Record<string, unknown> = { ...entry };
  if (Array.isArray(entry.definitions)) {
    pruned.definitions = entry.definitions.map(pruneToDictionaryNameFields);
  }
  if (Array.isArray(entry.pronunciations)) {
    pruned.pronunciations = entry.pronunciations.map(pruneToDictionaryNameFields);
  }
  if (Array.isArray(entry.frequencies)) {
    pruned.frequencies = entry.frequencies.map(pruneFrequencyItem);
  }
  return pruned;
}

function pruneMessageResult(action: string, result: unknown): unknown {
  if (action === 'optionsGetFull' && isRecord(result)) {
    const profiles = Array.isArray(result.profiles) ? result.profiles : [];
    return {
      profileCurrent: result.profileCurrent,
      profiles: profiles.map((profile) => {
        if (!isRecord(profile) || !isRecord(profile.options)) {
          return {};
        }
        const options = profile.options;
        return {
          options: {
            scanning: isRecord(options.scanning) ? { length: options.scanning.length } : {},
            dictionaries: Array.isArray(options.dictionaries)
              ? options.dictionaries.map((dictionary) =>
                  isRecord(dictionary)
                    ? {
                        name: dictionary.name,
                        enabled: dictionary.enabled,
                        id: dictionary.id,
                      }
                    : {},
                )
              : [],
          },
        };
      }),
    };
  }
  if (action === 'getDictionaryInfo' && Array.isArray(result)) {
    return result.map((entry) =>
      isRecord(entry) ? { title: entry.title, frequencyMode: entry.frequencyMode } : {},
    );
  }
  if (action === 'termsFind' && isRecord(result) && Array.isArray(result.dictionaryEntries)) {
    return {
      ...result,
      dictionaryEntries: result.dictionaryEntries.map(pruneTermsFindEntry),
    };
  }
  return result;
}

// Strip recorded page traffic down to what replay can ever request: drop
// Yomitan's own background messages and the heavyweight response fields
// (definition glossaries, full options blobs) the injected helpers never read.
// Prune safety is verified by the replay tests themselves — a pruned fixture
// must still reproduce its recorded expectation.
export function pruneGoldenMessages(messages: GoldenRecordedMessage[]): GoldenRecordedMessage[] {
  return messages
    .filter((message) => GOLDEN_MESSAGE_ACTIONS.has(message.action))
    .map((message) => {
      if (isRecord(message.response) && 'result' in message.response) {
        return {
          ...message,
          response: {
            ...message.response,
            result: pruneMessageResult(message.action, message.response.result),
          },
        };
      }
      return message;
    });
}

export function hashInjectedScript(script: string): string {
  return createHash('sha256').update(script).digest('hex');
}

export function classifyInjectedScript(script: string): string {
  const markers = [
    'optionsGetFull',
    'parseText',
    'termsFind',
    'getTermFrequencies',
    'setAllSettings',
  ];
  const found = markers.filter((marker) => script.includes(marker));
  return found.length > 0 ? found.join('+') : 'unknown';
}

function stableStringify(value: unknown): string {
  // Recorded params pass through JSON, which turns undefined into null.
  return JSON.stringify(value ?? null, (_key, entry: unknown) => {
    if (entry && typeof entry === 'object' && !Array.isArray(entry)) {
      const record = entry as Record<string, unknown>;
      const sorted: Record<string, unknown> = {};
      for (const key of Object.keys(record).sort()) {
        sorted[key] = record[key];
      }
      return sorted;
    }
    return entry;
  });
}

export interface FixtureLookups {
  isKnownWord: (
    text: string,
    reading?: string,
    options?: { allowReadingOnlyMatch?: boolean },
  ) => boolean;
  getJlptLevel: (text: string) => JlptLevel | null;
  getFrequencyRank: (term: string) => number | null;
}

// Deliberately simple known-word semantics (exact text, optional reading
// constraint) shared by the recorder and the replay harness so both sides of
// a fixture agree. The production reading-aware cache is exercised by its own
// unit tests, not by the golden corpus.
export function buildFixtureLookups(config: GoldenFixtureConfig): FixtureLookups {
  const readingsByText = new Map<string, Set<string> | null>();
  for (const entry of config.knownWords ?? []) {
    const text = typeof entry === 'string' ? entry : entry.text;
    const reading = typeof entry === 'string' ? undefined : entry.reading;
    if (!reading) {
      readingsByText.set(text, null);
      continue;
    }
    const existing = readingsByText.get(text);
    if (existing === null) {
      continue;
    }
    const readings = existing ?? new Set<string>();
    readings.add(reading);
    readingsByText.set(text, readings);
  }

  return {
    isKnownWord: (text, reading) => {
      if (!readingsByText.has(text)) {
        return false;
      }
      const readings = readingsByText.get(text);
      if (readings === null || readings === undefined) {
        return true;
      }
      return reading === undefined || readings.has(reading);
    },
    getJlptLevel: (text) => config.jlptLevels?.[text] ?? null,
    getFrequencyRank: (term) => config.localFrequencyRanks?.[term] ?? null,
  };
}

// Config-driven deps getters shared by the recorder and the replay harness so
// a fixture is tokenized under identical toggles in both directions.
export function fixtureConfigGetters(config: GoldenFixtureConfig) {
  return {
    getKnownWordMatchMode: () => config.knownWordMatchMode ?? ('headword' as NPlusOneMatchMode),
    getKnownWordsEnabled: () => config.knownWordsEnabled !== false,
    getNPlusOneEnabled: () => config.nPlusOneEnabled !== false,
    getJlptEnabled: () => config.jlptEnabled !== false,
    getNameMatchEnabled: () => config.nameMatchEnabled !== false,
    getNameMatchImagesEnabled: () => false,
    getFrequencyDictionaryEnabled: () => config.frequencyEnabled !== false,
    getFrequencyDictionaryMatchMode: () =>
      config.frequencyMatchMode ?? ('headword' as FrequencyDictionaryMatchMode),
    getMinSentenceWordsForNPlusOne: () => config.minSentenceWordsForNPlusOne ?? 3,
  };
}

interface ReplayMessageEntry extends GoldenRecordedMessage {
  used: boolean;
  paramsKey: string;
}

export interface ReplayMessageStore {
  handle: (action: string, params: unknown) => unknown;
  unusedCount: () => number;
}

export function createReplayMessageStore(messages: GoldenRecordedMessage[]): ReplayMessageStore {
  const entries: ReplayMessageEntry[] = messages.map((message) => ({
    ...message,
    used: false,
    paramsKey: stableStringify(message.params),
  }));

  return {
    handle: (action, params) => {
      const paramsKey = stableStringify(params);
      const match =
        entries.find((e) => !e.used && e.action === action && e.paramsKey === paramsKey) ??
        entries.find((e) => !e.used && e.action === action) ??
        entries.find((e) => e.action === action && e.paramsKey === paramsKey);
      if (!match) {
        throw new Error(
          `golden fixture replay: no recorded response for action "${action}" params=${paramsKey}`,
        );
      }
      match.used = true;
      return match.response;
    },
    unusedCount: () => entries.filter((e) => !e.used).length,
  };
}

// One persistent context per fixture, matching the real parser window: the
// scan runtime installs itself once into globalThis and later per-line call
// scripts reuse it.
function createInjectedScriptVm(store: ReplayMessageStore): (script: string) => Promise<unknown> {
  const context = vm.createContext({
    chrome: {
      runtime: {
        lastError: null,
        sendMessage: (
          payload: { action?: string; params?: unknown },
          callback: (response: unknown) => void,
        ) => {
          callback(store.handle(payload.action ?? '', payload.params));
        },
      },
    },
    Array,
    Boolean,
    Date,
    Error,
    JSON,
    Map,
    Math,
    Number,
    Object,
    Promise,
    RegExp,
    Set,
    String,
  });
  return async (script: string) => await vm.runInContext(script, context);
}

export function createReplayTokenizerDeps(fixture: GoldenFixture): TokenizerServiceDeps {
  const store = createReplayMessageStore(fixture.recording.messages);
  const scriptResults = new Map(
    fixture.recording.scripts.map((entry) => [entry.sha256, entry] as const),
  );
  const runInjectedScriptInVm = createInjectedScriptVm(store);

  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async (script: string) => {
        try {
          return await runInjectedScriptInVm(script);
        } catch (vmError) {
          const recorded = scriptResults.get(hashInjectedScript(script));
          if (recorded) {
            return recorded.result;
          }
          throw new Error(
            `golden fixture "${fixture.name}": injected script (${classifyInjectedScript(script)}) failed in vm replay and has no recorded script-level result: ${(vmError as Error).message}`,
          );
        }
      },
    },
  };

  const lookups = buildFixtureLookups(fixture.config);
  const deps = createTokenizerDepsRuntime({
    getYomitanExt: () => ({ id: 'golden-fixture-extension' }) as never,
    getYomitanParserWindow: () => parserWindow as never,
    setYomitanParserWindow: () => undefined,
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => undefined,
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => undefined,
    isKnownWord: lookups.isKnownWord,
    getJlptLevel: lookups.getJlptLevel,
    getFrequencyRank: lookups.getFrequencyRank,
    ...fixtureConfigGetters(fixture.config),
    getMecabTokenizer: () => ({
      tokenize: async (text: string) => fixture.recording.mecab[text] ?? null,
    }),
  });

  return {
    ...deps,
    // Pin the synchronous enrichment stage so record and replay share one
    // implementation and no worker spawns inside bun test.
    enrichTokensWithMecab: async (tokens, mecabTokens) =>
      enrichTokensWithMecabPos1(tokens, mecabTokens),
  };
}

export function projectGoldenTokens(tokens: MergedToken[] | null): GoldenExpectedToken[] | null {
  if (!tokens) {
    return null;
  }
  return tokens.map((token) => {
    const projected: GoldenExpectedToken = {
      surface: token.surface,
      reading: token.reading,
      headword: token.headword,
      startPos: token.startPos,
      endPos: token.endPos,
      partOfSpeech: token.partOfSpeech,
      isKnown: token.isKnown,
      isNPlusOneTarget: token.isNPlusOneTarget,
    };
    if (token.headwordReading !== undefined) projected.headwordReading = token.headwordReading;
    if (token.pos1 !== undefined) projected.pos1 = token.pos1;
    if (token.pos2 !== undefined) projected.pos2 = token.pos2;
    if (token.pos3 !== undefined) projected.pos3 = token.pos3;
    if (token.isNameMatch === true) projected.isNameMatch = true;
    if (token.isUnparsedRun === true) projected.isUnparsedRun = true;
    if (token.jlptLevel !== undefined) projected.jlptLevel = token.jlptLevel;
    if (token.frequencyRank !== undefined) projected.frequencyRank = token.frequencyRank;
    return projected;
  });
}
