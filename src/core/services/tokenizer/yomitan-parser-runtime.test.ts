import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import * as vm from 'node:vm';
import {
  addYomitanNoteViaSearch,
  clearYomitanParserCachesForWindow,
  extractYomitanCurrentAnkiDeckName,
  getYomitanDictionaryInfo,
  importYomitanDictionaryFromZip,
  deleteYomitanDictionaryByTitle,
  removeYomitanDictionarySettings,
  requestYomitanScanTokens,
  requestYomitanTermFrequencies,
  syncYomitanDefaultAnkiServer,
  upsertYomitanDictionarySettings,
} from './yomitan-parser-runtime';

function createDeps(
  executeJavaScript: (script: string) => Promise<unknown>,
  options?: {
    createYomitanExtensionWindow?: (pageName: string) => Promise<unknown>;
  },
) {
  const parserWindow = {
    isDestroyed: () => false,
    webContents: {
      executeJavaScript: async (script: string) => await executeJavaScript(script),
    },
  };

  return {
    getYomitanExt: () => ({ id: 'ext-id' }) as never,
    getYomitanParserWindow: () => parserWindow as never,
    setYomitanParserWindow: () => undefined,
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => undefined,
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => undefined,
    createYomitanExtensionWindow: options?.createYomitanExtensionWindow as never,
  };
}

function createYomitanScriptSandbox(handler: (action: string, params: unknown) => unknown) {
  return {
    chrome: {
      runtime: {
        lastError: null,
        sendMessage: (
          payload: { action?: string; params?: unknown },
          callback: (response: { result?: unknown; error?: { message?: string } }) => void,
        ) => {
          try {
            callback({ result: handler(payload.action ?? '', payload.params) });
          } catch (error) {
            callback({ error: { message: (error as Error).message } });
          }
        },
      },
    },
    Array,
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
  };
}

async function runInjectedYomitanScript(
  script: string,
  handler: (action: string, params: unknown) => unknown,
): Promise<unknown> {
  return await vm.runInNewContext(script, createYomitanScriptSandbox(handler));
}

// Persistent page context shared across executeJavaScript calls, matching the
// real parser window: the scan runtime is installed once via
// globalThis.__subminerYomitanScan and per-line calls reuse it (and its
// cross-line termsFind cache).
function createPersistentYomitanScriptRunner(
  handler: (action: string, params: unknown) => unknown,
): (script: string) => Promise<unknown> {
  const context = vm.createContext(createYomitanScriptSandbox(handler));
  return async (script: string) => await vm.runInContext(script, context);
}

// Deps whose parser window executes every injected script (profile metadata,
// scan runtime install, per-line scan calls, parseText fallback) inside one
// persistent vm context, dispatching backend actions to `handler`.
function createScanDeps(
  handler: (action: string, params: unknown) => unknown,
  options?: { onScript?: (script: string) => void },
) {
  const runScript = createPersistentYomitanScriptRunner(handler);
  return createDeps(async (script) => {
    options?.onScript?.(script);
    return await runScript(script);
  });
}

function countTermsFindLookups(lookups: string[], prefix: string): number {
  return lookups.filter((lookupText) => lookupText.startsWith(prefix)).length;
}

// Backend stub for the greedy name pre-pass: one character name (ミナト) in a
// line of ordinary words, with the SubMiner character dictionary enabled.
const NAME_SCAN_WORDS: Array<[string, string, string, boolean]> = [
  ['ミナト', 'ミナト', 'みなと', true],
  ['は', 'は', 'は', false],
  ['まだ', 'まだ', 'まだ', false],
  ['学校', '学校', 'がっこう', false],
  ['に', 'に', 'に', false],
  ['いない', 'いる', 'いる', false],
];

function createNameScanDeps(
  lookups: string[],
  words: Array<[string, string, string, boolean]> = NAME_SCAN_WORDS,
) {
  return createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JMdict', enabled: true, id: 0 },
                {
                  name: 'SubMiner Character Dictionary (AniList 1)',
                  enabled: true,
                  id: 1,
                },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    lookups.push(text);
    for (const [surface, term, reading, isName] of words) {
      if (text.startsWith(surface)) {
        return {
          originalTextLength: surface.length,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term,
                  reading,
                  sources: [{ originalText: surface, isPrimary: true, matchType: 'exact' }],
                },
              ],
              definitions: [
                { dictionary: isName ? 'SubMiner Character Dictionary (AniList 1)' : 'JMdict' },
              ],
            },
          ],
        };
      }
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });
}

const NAME_SCAN_LINE = 'ミナトはまだ学校にいない';

test('requestYomitanScanTokens skips name pre-pass lookups where no candidate name can start', async () => {
  const exhaustiveLookups: string[] = [];
  const exhaustive = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(exhaustiveLookups),
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  const prefilteredLookups: string[] = [];
  const prefiltered = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(prefilteredLookups),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      // Terms and readings the generated dictionary exposes for this media.
      nameCandidates: { key: 'media-1', forms: ['ミナト', 'みなと'] },
    },
  );

  // Same tokenization, including the name match, with fewer round trips.
  assert.deepEqual(prefiltered, exhaustive);
  assert.equal(prefiltered?.[0]?.surface, 'ミナト');
  assert.equal(prefiltered?.[0]?.isNameMatch, true);
  assert.ok(
    prefilteredLookups.length < exhaustiveLookups.length,
    `expected fewer lookups with candidates (${prefilteredLookups.length} vs ${exhaustiveLookups.length})`,
  );
  // Mid-token positions are exactly what the pre-pass used to probe (a name can
  // start mid-token); with candidates they cost nothing, while the main walk's
  // own token-start lookups are unaffected.
  assert.ok(countTermsFindLookups(exhaustiveLookups, '校に') > 0);
  assert.equal(countTermsFindLookups(prefilteredLookups, '校に'), 0);
});

test('requestYomitanScanTokens matches a katakana name from its kana-normalized candidate form', async () => {
  const lookups: string[] = [];
  const result = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(lookups),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      // Only the hiragana reading is listed; the katakana surface in the line
      // must still be found through kana normalization.
      nameCandidates: { key: 'media-1', forms: ['みなと'] },
    },
  );

  assert.equal(result?.[0]?.surface, 'ミナト');
  assert.equal(result?.[0]?.isNameMatch, true);
});

// Kana normalization does not fold halfwidth katakana, so a name written that
// way can never prefix-match a candidate form; the pre-pass has a bypass for
// those positions, which only runs if they count as Japanese in the first place.
// The generic word here reaches into the name, so only a pre-pass reservation
// can keep the name whole.
const HALFWIDTH_NAME_SCAN_WORDS: Array<[string, string, string, boolean]> = [
  ['まだﾐ', 'まだミ', 'まだみ', false],
  ['まだ', 'まだ', 'まだ', false],
  ['ﾐﾅﾄ', 'ミナト', 'みなと', true],
];

test('requestYomitanScanTokens probes halfwidth katakana positions during the name pre-pass', async () => {
  const lookups: string[] = [];
  const result = await requestYomitanScanTokens(
    'まだﾐﾅﾄ',
    createNameScanDeps(lookups, HALFWIDTH_NAME_SCAN_WORDS),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      // Fullwidth forms only, as the generated dictionary stores them.
      nameCandidates: { key: 'media-1', forms: ['ミナト', 'みなと'] },
    },
  );

  assert.equal(countTermsFindLookups(lookups, 'ﾐﾅﾄ'), 1);
  assert.deepEqual(
    result?.map((token) => token.surface),
    ['まだ', 'ﾐﾅﾄ'],
  );
  assert.equal(result?.[1]?.isNameMatch, true);
});

test('requestYomitanScanTokens falls back to the exhaustive name scan without candidates', async () => {
  const withoutLookups: string[] = [];
  const withoutCandidates = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(withoutLookups),
    { error: () => undefined },
    { includeNameMatchMetadata: true, currentCharacterDictionaryMediaId: 1, nameCandidates: null },
  );

  assert.equal(withoutCandidates?.[0]?.isNameMatch, true);
  // No candidate list means every Japanese position is probed, as before.
  assert.ok(countTermsFindLookups(withoutLookups, '校に') > 0);
});

test('requestYomitanScanTokens reinstalls name candidates when the media changes', async () => {
  const lookups: string[] = [];
  const deps = createNameScanDeps(lookups);

  // First media's candidates cannot match this line's name.
  const otherMedia = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    deps,
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 2,
      nameCandidates: { key: 'media-2', forms: ['カズマ'] },
    },
  );
  assert.equal(otherMedia?.[0]?.isNameMatch, undefined);

  const correctMedia = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    deps,
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      nameCandidates: { key: 'media-1', forms: ['ミナト'] },
    },
  );
  assert.equal(correctMedia?.[0]?.surface, 'ミナト');
  assert.equal(correctMedia?.[0]?.isNameMatch, true);
});

test('syncYomitanDefaultAnkiServer updates default profile server when script reports update', async () => {
  let scriptValue = '';
  const deps = createDeps(async (script) => {
    scriptValue = script;
    return { updated: true };
  });

  const infoLogs: string[] = [];
  const updated = await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', deps, {
    error: () => undefined,
    info: (message) => infoLogs.push(message),
  });

  assert.equal(updated, true);
  assert.match(scriptValue, /optionsGetFull/);
  assert.match(scriptValue, /setAllSettings/);
  assert.match(scriptValue, /profileCurrent/);
  assert.match(scriptValue, /forceOverride = false/);
  assert.equal(infoLogs.length, 1);
});

test('syncYomitanDefaultAnkiServer returns true when script reports no change', async () => {
  const deps = createDeps(async () => ({ updated: false }));
  let infoLogCount = 0;

  const synced = await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', deps, {
    error: () => undefined,
    info: () => {
      infoLogCount += 1;
    },
  });

  assert.equal(synced, true);
  assert.equal(infoLogCount, 0);
});

test('syncYomitanDefaultAnkiServer returns false when existing non-default server blocks update', async () => {
  const deps = createDeps(async () => ({
    updated: false,
    matched: false,
    reason: 'blocked-existing-server',
  }));
  const infoLogs: string[] = [];

  const synced = await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', deps, {
    error: () => undefined,
    info: (message) => infoLogs.push(message),
  });

  assert.equal(synced, false);
  assert.equal(infoLogs.length, 1);
  assert.match(infoLogs[0] ?? '', /blocked-existing-server/);
});

test('syncYomitanDefaultAnkiServer injects force override when enabled', async () => {
  let scriptValue = '';
  const deps = createDeps(async (script) => {
    scriptValue = script;
    return { updated: false, matched: true };
  });

  const synced = await syncYomitanDefaultAnkiServer(
    'http://127.0.0.1:8766',
    deps,
    {
      error: () => undefined,
      info: () => undefined,
    },
    { forceOverride: true },
  );

  assert.equal(synced, true);
  assert.match(scriptValue, /forceOverride = true/);
});

test('syncYomitanDefaultAnkiServer updates the active profile Anki deck', async () => {
  const optionsFull = {
    profileCurrent: 0,
    profiles: [
      {
        options: {
          anki: {
            server: 'http://127.0.0.1:8766',
            cardFormats: [
              { type: 'term', deck: 'Default', model: 'Mining Note', fields: {} },
              { type: 'kanji', deck: 'Kanji', model: 'Kanji Note', fields: {} },
            ],
            terms: { deck: 'Default', model: 'Legacy Note', fields: {} },
          },
        },
      },
    ],
  };
  let savedOptions: typeof optionsFull | null = null;
  const deps = createDeps((script) =>
    runInjectedYomitanScript(script, (action, params) => {
      if (action === 'optionsGetFull') {
        return JSON.parse(JSON.stringify(optionsFull));
      }
      if (action === 'setAllSettings') {
        savedOptions = (params as { value: typeof optionsFull }).value;
        return true;
      }
      throw new Error(`Unexpected action: ${action}`);
    }),
  );

  const synced = await syncYomitanDefaultAnkiServer(
    'http://127.0.0.1:8766',
    deps,
    {
      error: () => undefined,
      info: () => undefined,
    },
    { deck: 'Minecraft', forceOverride: true },
  );

  assert.equal(synced, true);
  assert.ok(savedOptions);
  const saved = savedOptions as typeof optionsFull;
  assert.equal(saved.profiles[0]?.options.anki.cardFormats[0]?.deck, 'Minecraft');
  assert.equal(saved.profiles[0]?.options.anki.cardFormats[1]?.deck, 'Kanji');
  assert.equal(saved.profiles[0]?.options.anki.terms.deck, 'Minecraft');
});

test('syncYomitanDefaultAnkiServer logs and returns false on script failure', async () => {
  const deps = createDeps(async () => {
    throw new Error('execute failed');
  });

  const errorLogs: string[] = [];
  const updated = await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', deps, {
    error: (message) => errorLogs.push(message),
    info: () => undefined,
  });

  assert.equal(updated, false);
  assert.equal(errorLogs.length, 1);
});

test('syncYomitanDefaultAnkiServer no-ops for empty target url', async () => {
  let executeCount = 0;
  const deps = createDeps(async () => {
    executeCount += 1;
    return { updated: true };
  });

  const updated = await syncYomitanDefaultAnkiServer('   ', deps, {
    error: () => undefined,
    info: () => undefined,
  });

  assert.equal(updated, false);
  assert.equal(executeCount, 0);
});

test('extractYomitanCurrentAnkiDeckName prefers the active profile first term card format deck', () => {
  assert.equal(
    extractYomitanCurrentAnkiDeckName({
      profileCurrent: 1,
      profiles: [
        {
          options: {
            anki: {
              cardFormats: [{ type: 'term', deck: 'Inactive' }],
            },
          },
        },
        {
          options: {
            anki: {
              cardFormats: [
                { type: 'kanji', deck: 'Kanji' },
                { type: 'term', deck: 'Mining' },
              ],
            },
          },
        },
      ],
    }),
    'Mining',
  );
});

test('extractYomitanCurrentAnkiDeckName ignores disabled card format decks', () => {
  assert.equal(
    extractYomitanCurrentAnkiDeckName({
      profiles: [
        {
          options: {
            anki: {
              cardFormats: [
                { type: 'term', deck: 'Disabled Term', enabled: false },
                { type: 'kanji', deck: 'Disabled Kanji', enabled: false },
                { type: 'term', deck: 'Mining', enabled: true },
              ],
            },
          },
        },
      ],
    }),
    'Mining',
  );
});

test('extractYomitanCurrentAnkiDeckName falls back to legacy term deck', () => {
  assert.equal(
    extractYomitanCurrentAnkiDeckName({
      profiles: [
        {
          options: {
            anki: {
              terms: { deck: 'Legacy Mining' },
            },
          },
        },
      ],
    }),
    'Legacy Mining',
  );
});

test('requestYomitanTermFrequencies returns normalized frequency entries', async () => {
  let scriptValue = '';
  const deps = createDeps(async (script) => {
    scriptValue = script;
    return [
      {
        term: '猫',
        reading: 'ねこ',
        hasReading: true,
        dictionary: 'freq-dict',
        dictionaryPriority: 0,
        frequency: 77,
        displayValue: '77',
        displayValueParsed: true,
      },
      {
        term: '鍛える',
        reading: 'きたえる',
        hasReading: false,
        dictionary: 'freq-dict',
        dictionaryPriority: 1,
        frequency: 46961,
        displayValue: '2847,46961',
        displayValueParsed: true,
      },
      {
        term: 'invalid',
        dictionary: 'freq-dict',
        frequency: 0,
      },
    ];
  });

  const result = await requestYomitanTermFrequencies([{ term: '猫', reading: 'ねこ' }], deps, {
    error: () => undefined,
  });

  assert.equal(result.length, 2);
  assert.equal(result[0]?.term, '猫');
  assert.equal(result[0]?.hasReading, true);
  assert.equal(result[0]?.frequency, 77);
  assert.equal(result[0]?.dictionaryPriority, 0);
  assert.equal(result[1]?.term, '鍛える');
  assert.equal(result[1]?.hasReading, false);
  assert.equal(result[1]?.frequency, 2847);
  assert.match(scriptValue, /getTermFrequencies/);
  assert.match(scriptValue, /optionsGetFull/);
});

test('requestYomitanTermFrequencies prefers primary rank from displayValue array pair', async () => {
  const deps = createDeps(async () => [
    {
      term: '無人',
      reading: 'むじん',
      dictionary: 'freq-dict',
      dictionaryPriority: 0,
      frequency: 157632,
      displayValue: [7141, 157632],
      displayValueParsed: true,
    },
  ]);

  const result = await requestYomitanTermFrequencies([{ term: '無人', reading: 'むじん' }], deps, {
    error: () => undefined,
  });

  assert.equal(result.length, 1);
  assert.equal(result[0]?.term, '無人');
  assert.equal(result[0]?.frequency, 7141);
});

test('requestYomitanTermFrequencies prefers primary rank from displayValue string pair when raw frequency matches trailing count', async () => {
  const deps = createDeps(async () => [
    {
      term: '潜む',
      reading: 'ひそむ',
      dictionary: 'freq-dict',
      dictionaryPriority: 0,
      frequency: 121,
      displayValue: '118,121',
      displayValueParsed: false,
    },
  ]);

  const result = await requestYomitanTermFrequencies([{ term: '潜む', reading: 'ひそむ' }], deps, {
    error: () => undefined,
  });

  assert.equal(result.length, 1);
  assert.equal(result[0]?.term, '潜む');
  assert.equal(result[0]?.frequency, 118);
});

test('requestYomitanTermFrequencies uses leading display digits for displayValue strings', async () => {
  const deps = createDeps(async () => [
    {
      term: '例',
      reading: 'れい',
      dictionary: 'freq-dict',
      dictionaryPriority: 0,
      frequency: 1234,
      displayValue: '1,234',
      displayValueParsed: false,
    },
  ]);

  const result = await requestYomitanTermFrequencies([{ term: '例', reading: 'れい' }], deps, {
    error: () => undefined,
  });

  assert.equal(result.length, 1);
  assert.equal(result[0]?.term, '例');
  assert.equal(result[0]?.frequency, 1);
});

test('requestYomitanTermFrequencies ignores occurrence-based dictionaries for rank tagging', async () => {
  let metadataScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('getTermFrequencies')) {
      return [
        {
          term: '潜む',
          reading: 'ひそむ',
          dictionary: 'CC100',
          frequency: 118121,
          displayValue: null,
          displayValueParsed: false,
        },
      ];
    }

    if (script.includes('optionsGetFull')) {
      metadataScript = script;
      return {
        profileCurrent: 0,
        profileIndex: 0,
        scanLength: 40,
        dictionaries: ['CC100'],
        dictionaryPriorityByName: { CC100: 0 },
        dictionaryFrequencyModeByName: { CC100: 'occurrence-based' },
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [{ name: 'CC100', enabled: true, id: 0 }],
            },
          },
        ],
      };
    }
    return [];
  });

  const result = await requestYomitanTermFrequencies([{ term: '潜む', reading: 'ひそむ' }], deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, []);
  assert.match(metadataScript, /getDictionaryInfo/);
});

test('requestYomitanTermFrequencies requests term-only fallback only after reading miss', async () => {
  const frequencyScripts: string[] = [];
  const deps = createDeps(async (script) => {
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [{ name: 'freq-dict', enabled: true, id: 0 }],
            },
          },
        ],
      };
    }

    if (!script.includes('getTermFrequencies')) {
      return [];
    }

    frequencyScripts.push(script);
    if (script.includes('"term":"断じて","reading":"だん"')) {
      return [];
    }
    if (script.includes('"term":"断じて","reading":null')) {
      return [
        {
          term: '断じて',
          reading: null,
          dictionary: 'freq-dict',
          frequency: 7082,
          displayValue: '7082',
          displayValueParsed: true,
        },
      ];
    }
    return [];
  });

  const result = await requestYomitanTermFrequencies([{ term: '断じて', reading: 'だん' }], deps, {
    error: () => undefined,
  });

  assert.equal(result.length, 1);
  assert.equal(result[0]?.frequency, 7082);
  assert.equal(frequencyScripts.length, 2);
  assert.match(frequencyScripts[0] ?? '', /"term":"断じて","reading":"だん"/);
  assert.doesNotMatch(frequencyScripts[0] ?? '', /"term":"断じて","reading":null/);
  assert.match(frequencyScripts[1] ?? '', /"term":"断じて","reading":null/);
});

test('requestYomitanTermFrequencies avoids term-only fallback request when reading lookup succeeds', async () => {
  const frequencyScripts: string[] = [];
  const deps = createDeps(async (script) => {
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [{ name: 'freq-dict', enabled: true, id: 0 }],
            },
          },
        ],
      };
    }

    if (!script.includes('getTermFrequencies')) {
      return [];
    }

    frequencyScripts.push(script);
    return [
      {
        term: '鍛える',
        reading: 'きたえる',
        dictionary: 'freq-dict',
        frequency: 2847,
        displayValue: '2847',
        displayValueParsed: true,
      },
    ];
  });

  const result = await requestYomitanTermFrequencies([{ term: '鍛える', reading: 'きた' }], deps, {
    error: () => undefined,
  });

  assert.equal(result.length, 1);
  assert.equal(frequencyScripts.length, 1);
  assert.match(frequencyScripts[0] ?? '', /"term":"鍛える","reading":"きた"/);
  assert.doesNotMatch(frequencyScripts[0] ?? '', /"term":"鍛える","reading":null/);
});

test('requestYomitanTermFrequencies caches profile metadata between calls', async () => {
  const scripts: string[] = [];
  const deps = createDeps(async (script) => {
    scripts.push(script);
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [{ name: 'freq-dict', enabled: true, id: 0 }],
            },
          },
        ],
      };
    }

    if (script.includes('"term":"犬"')) {
      return [
        {
          term: '犬',
          reading: 'いぬ',
          dictionary: 'freq-dict',
          frequency: 12,
          displayValue: '12',
          displayValueParsed: true,
        },
      ];
    }

    return [
      {
        term: '猫',
        reading: 'ねこ',
        dictionary: 'freq-dict',
        frequency: 77,
        displayValue: '77',
        displayValueParsed: true,
      },
    ];
  });

  await requestYomitanTermFrequencies([{ term: '猫', reading: 'ねこ' }], deps, {
    error: () => undefined,
  });
  await requestYomitanTermFrequencies([{ term: '犬', reading: 'いぬ' }], deps, {
    error: () => undefined,
  });

  const optionsCalls = scripts.filter((script) => script.includes('optionsGetFull')).length;
  assert.equal(optionsCalls, 1);
});

test('requestYomitanTermFrequencies caches repeated term+reading lookups', async () => {
  const scripts: string[] = [];
  const deps = createDeps(async (script) => {
    scripts.push(script);
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [{ name: 'freq-dict', enabled: true, id: 0 }],
            },
          },
        ],
      };
    }

    return [
      {
        term: '猫',
        reading: 'ねこ',
        dictionary: 'freq-dict',
        frequency: 77,
        displayValue: '77',
        displayValueParsed: true,
      },
    ];
  });

  await requestYomitanTermFrequencies([{ term: '猫', reading: 'ねこ' }], deps, {
    error: () => undefined,
  });
  await requestYomitanTermFrequencies([{ term: '猫', reading: 'ねこ' }], deps, {
    error: () => undefined,
  });

  const frequencyCalls = scripts.filter((script) => script.includes('getTermFrequencies')).length;
  assert.equal(frequencyCalls, 1);
});

test('requestYomitanScanTokens tokenizes with the in-window scanner and no parseText request', async () => {
  const scripts: string[] = [];
  const actions: string[] = [];
  const deps = createScanDeps(
    (action, params) => {
      actions.push(action);
      if (action === 'optionsGetFull') {
        return {
          profileCurrent: 0,
          profiles: [{ options: { scanning: { length: 40 } } }],
        };
      }
      if (action === 'getDictionaryInfo') {
        return [];
      }
      if (action === 'termsFind') {
        const text = (params as { text?: string } | undefined)?.text ?? '';
        if (!text.startsWith('取り組んで')) {
          return { originalTextLength: 0, dictionaryEntries: [] };
        }
        return {
          originalTextLength: 5,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term: '取り組む',
                  reading: 'とりくむ',
                  sources: [{ originalText: '取り組んで', isPrimary: true, matchType: 'exact' }],
                },
              ],
            },
          ],
        };
      }
      throw new Error(`unexpected action: ${action}`);
    },
    { onScript: (script) => scripts.push(script) },
  );

  const result = await requestYomitanScanTokens('取り組んで', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, [
    {
      surface: '取り組んで',
      reading: 'とりくんで',
      headword: '取り組む',
      headwordReading: 'とりくむ',
      startPos: 0,
      endPos: 5,
      isNameMatch: false,
      frequencyRank: undefined,
    },
  ]);
  // The duplicate full parse per line is gone: the scanner walk is the only
  // tokenization request.
  assert.ok(!actions.includes('parseText'));
  const installScript = scripts.find((script) => script.includes('termsFind'));
  assert.ok(installScript, 'expected the scan runtime install script');
  assert.match(installScript ?? '', /matchType:\s*"exact"/);
  assert.match(installScript ?? '', /deinflect:\s*true/);
});

test('requestYomitanScanTokens warns when active Yomitan profile has no dictionaries', async () => {
  const warnings: Array<{ message: string; details: unknown }> = [];
  const deps = createScanDeps((action) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });

  await requestYomitanScanTokens('字幕', deps, {
    error: () => undefined,
    warn: (message, details) => warnings.push({ message, details }),
  });

  assert.equal(warnings.length, 1);
  assert.match(warnings[0]!.message, /no enabled dictionaries/);
  assert.deepEqual(warnings[0]!.details, {
    profileIndex: 0,
    scanLength: 40,
    dictionaryCount: 0,
    dictionaries: [],
    omittedDictionaryCount: 0,
  });
});

test('requestYomitanScanTokens keeps reading aligned when a kana run extends the previous token', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    // 待ち合わせ matches, the trailing る does not, so the kana run extends the
    // previous token instead of becoming its own filler token.
    if (text.startsWith('待ち合わせ')) {
      return {
        originalTextLength: 5,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: '待ち合わせる',
                reading: 'まちあわせる',
                sources: [{ originalText: '待ち合わせ', isPrimary: true, matchType: 'exact' }],
              },
            ],
          },
        ],
      };
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });

  const result = await requestYomitanScanTokens('待ち合わせる', deps, {
    error: () => undefined,
  });

  assert.equal(result?.length, 1);
  assert.equal(result?.[0]?.surface, '待ち合わせる');
  assert.equal(result?.[0]?.endPos, 6);
  // The reading must grow with the surface: a short reading fails
  // isCompleteReadingForSurface and silently disables the known-word reading
  // fallback downstream.
  assert.equal(result?.[0]?.reading, 'まちあわせる');
});

test('requestYomitanScanTokens emits unparsed filler runs for text the scanner skips', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    const singleCharEntry = (term: string, reading: string) => ({
      originalTextLength: 1,
      dictionaryEntries: [
        {
          headwords: [
            {
              term,
              reading,
              sources: [{ originalText: text[0], isPrimary: true, matchType: 'exact' }],
            },
          ],
        },
      ],
    });
    if (text.startsWith('や')) {
      return singleCharEntry('や', 'や');
    }
    if (text.startsWith('ほ')) {
      return singleCharEntry('帆', 'ほ');
    }
    if (text.startsWith('ミナト')) {
      return {
        originalTextLength: 3,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: 'ミナト',
                reading: 'みなと',
                sources: [{ originalText: 'ミナト', isPrimary: true, matchType: 'exact' }],
              },
            ],
          },
        ],
      };
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });

  const result = await requestYomitanScanTokens('やほっ ミナト', deps, {
    error: () => undefined,
  });

  assert.deepEqual(
    result?.map(({ surface, headword, startPos, endPos, isUnparsedRun }) => ({
      surface,
      headword,
      startPos,
      endPos,
      isUnparsedRun,
    })),
    [
      { surface: 'や', headword: 'や', startPos: 0, endPos: 1, isUnparsedRun: undefined },
      { surface: 'ほ', headword: '帆', startPos: 1, endPos: 2, isUnparsedRun: undefined },
      // The unmatched っ + space becomes a hoverable filler run, replacing the
      // parseText filler chunks the pipeline used to rely on.
      { surface: 'っ ', headword: 'っ ', startPos: 2, endPos: 4, isUnparsedRun: true },
      { surface: 'ミナト', headword: 'ミナト', startPos: 4, endPos: 7, isUnparsedRun: undefined },
    ],
  );
  assert.equal(result?.[2]?.reading, '');
});

test('requestYomitanScanTokens extracts best frequency rank from selected termsFind entry', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JPDBv2㋕', enabled: true, id: 0 },
                { name: 'Jiten', enabled: true, id: 1 },
                { name: 'CC100', enabled: true, id: 2 },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (!text.startsWith('潜み')) {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }

    return {
      originalTextLength: 2,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: '潜む',
              reading: 'ひそむ',
              sources: [{ originalText: '潜み', isPrimary: true, matchType: 'exact' }],
            },
          ],
          frequencies: [
            {
              headwordIndex: 0,
              dictionary: 'JPDBv2㋕',
              frequency: 20181,
              displayValue: '4073,20181句',
            },
            {
              headwordIndex: 0,
              dictionary: 'Jiten',
              frequency: 28594,
              displayValue: '4592,28594句',
            },
            {
              headwordIndex: 0,
              dictionary: 'CC100',
              frequency: 118121,
              displayValue: null,
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('潜み', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, [
    {
      surface: '潜み',
      reading: 'ひそみ',
      headword: '潜む',
      headwordReading: 'ひそむ',
      startPos: 0,
      endPos: 2,
      isNameMatch: false,
      frequencyRank: 4073,
    },
  ]);
});

test('requestYomitanScanTokens retries shorter windows when a greedy match has no exact-source headword', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [{ name: 'JMdict', enabled: true, id: 0 }],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (!text.startsWith('平')) {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }
    if (text.length >= 4) {
      // Simulates Yomitan normalization consuming punctuation/whitespace:
      // the greedy match spans 平 （平 but no headword source equals it.
      return {
        originalTextLength: 4,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: '平々',
                reading: 'へいへい',
                sources: [{ originalText: '平平', isPrimary: true, matchType: 'exact' }],
              },
            ],
          },
        ],
      };
    }
    return {
      originalTextLength: 1,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: '平',
              reading: 'たいら',
              sources: [{ originalText: '平', isPrimary: true, matchType: 'exact' }],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('平 （平）', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, [
    {
      surface: '平',
      reading: 'たいら',
      headword: '平',
      headwordReading: 'たいら',
      startPos: 0,
      endPos: 1,
      isNameMatch: false,
      frequencyRank: undefined,
    },
    {
      surface: '平',
      reading: 'たいら',
      headword: '平',
      headwordReading: 'たいら',
      startPos: 3,
      endPos: 4,
      isNameMatch: false,
      frequencyRank: undefined,
    },
  ]);
});

test('requestYomitanScanTokens emits complete readings for kanji-kana compounds', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [{ name: 'JPDBv2㋕', enabled: true, id: 0 }],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (!text.startsWith('待ち合わせてる')) {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }

    return {
      originalTextLength: 7,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: '待ち合わせる',
              reading: 'まちあわせる',
              sources: [{ originalText: '待ち合わせてる', isPrimary: true, matchType: 'exact' }],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('待ち合わせてる', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, [
    {
      surface: '待ち合わせてる',
      reading: 'まちあわせてる',
      headword: '待ち合わせる',
      headwordReading: 'まちあわせる',
      startPos: 0,
      endPos: 7,
      isNameMatch: false,
      frequencyRank: undefined,
    },
  ]);
});

test('requestYomitanScanTokens uses frequency from later exact-match entry when first exact entry has none', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JPDBv2㋕', enabled: true, id: 0 },
                { name: 'Jiten', enabled: true, id: 1 },
                { name: 'CC100', enabled: true, id: 2 },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (!text.startsWith('者')) {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }

    return {
      originalTextLength: 1,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: '者',
              reading: 'もの',
              sources: [{ originalText: '者', isPrimary: true, matchType: 'exact' }],
            },
          ],
          frequencies: [],
        },
        {
          headwords: [
            {
              term: '者',
              reading: 'もの',
              sources: [{ originalText: '者', isPrimary: true, matchType: 'exact' }],
            },
          ],
          frequencies: [
            {
              headwordIndex: 0,
              dictionary: 'JPDBv2㋕',
              frequency: 79601,
              displayValue: '475,79601句',
            },
            {
              headwordIndex: 0,
              dictionary: 'Jiten',
              frequency: 338,
              displayValue: '338',
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('者', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, [
    {
      surface: '者',
      reading: 'もの',
      headword: '者',
      headwordReading: 'もの',
      startPos: 0,
      endPos: 1,
      isNameMatch: false,
      frequencyRank: 475,
    },
  ]);
});

test('requestYomitanScanTokens can use frequency from later exact secondary-match entry', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JPDBv2㋕', enabled: true, id: 0 },
                { name: 'Jiten', enabled: true, id: 1 },
                { name: 'CC100', enabled: true, id: 2 },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (!text.startsWith('者')) {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }

    return {
      originalTextLength: 1,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: '者',
              reading: 'もの',
              sources: [{ originalText: '者', isPrimary: true, matchType: 'exact' }],
            },
          ],
          frequencies: [],
        },
        {
          headwords: [
            {
              term: '者',
              reading: 'もの',
              sources: [{ originalText: '者', isPrimary: false, matchType: 'exact' }],
            },
          ],
          frequencies: [
            {
              headwordIndex: 0,
              dictionary: 'JPDBv2㋕',
              frequency: 79601,
              displayValue: '475,79601句',
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('者', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, [
    {
      surface: '者',
      reading: 'もの',
      headword: '者',
      headwordReading: 'もの',
      startPos: 0,
      endPos: 1,
      isNameMatch: false,
      frequencyRank: 475,
    },
  ]);
});

test('requestYomitanScanTokens uses exact frequency entry when selected reading differs', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JPDBv2㋕', enabled: true, id: 0 },
                { name: 'Jiten', enabled: true, id: 1 },
                { name: 'CC100', enabled: true, id: 2 },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (!text.startsWith('第二')) {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }

    return {
      originalTextLength: 2,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: '第二',
              reading: 'だいに',
              sources: [{ originalText: '第二', isPrimary: true, matchType: 'exact' }],
            },
          ],
          frequencies: [],
        },
        {
          headwords: [
            {
              term: '第二',
              reading: '',
              sources: [{ originalText: '第二', isPrimary: false, matchType: 'exact' }],
            },
          ],
          frequencies: [
            {
              headwordIndex: 0,
              dictionary: 'JPDBv2㋕',
              frequency: 189513,
              displayValue: '1820,189513句',
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('第二走者', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result?.[0], {
    surface: '第二',
    reading: 'だいに',
    headword: '第二',
    headwordReading: 'だいに',
    startPos: 0,
    endPos: 2,
    isNameMatch: false,
    frequencyRank: 1820,
  });
});

test('requestYomitanScanTokens marks tokens backed by SubMiner character dictionary entries', async () => {
  const deps = createDeps(async (script) => {
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
            },
          },
        ],
      };
    }

    return [
      {
        surface: 'アクア',
        reading: 'あくあ',
        headword: 'アクア',
        startPos: 0,
        endPos: 3,
        isNameMatch: true,
      },
      {
        surface: 'です',
        reading: 'です',
        headword: 'です',
        startPos: 3,
        endPos: 5,
        isNameMatch: false,
      },
    ];
  });

  const result = await requestYomitanScanTokens('アクアです', deps, {
    error: () => undefined,
  });

  assert.equal(result?.length, 2);
  assert.equal((result?.[0] as { isNameMatch?: boolean } | undefined)?.isNameMatch, true);
  assert.equal((result?.[1] as { isNameMatch?: boolean } | undefined)?.isNameMatch, false);
});

test('requestYomitanScanTokens skips name-match work when disabled', async () => {
  let scanCallScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('__subminerYomitanScan(')) {
      scanCallScript = script;
    }
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
            },
          },
        ],
      };
    }

    return [
      {
        surface: 'アクア',
        reading: 'あくあ',
        headword: 'アクア',
        startPos: 0,
        endPos: 3,
      },
    ];
  });

  const result = await requestYomitanScanTokens(
    'アクア',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: false },
  );

  assert.equal(result?.length, 1);
  assert.equal((result?.[0] as { isNameMatch?: boolean } | undefined)?.isNameMatch, undefined);
  assert.match(scanCallScript, /"includeNameMatchMetadata":false/);
});

test('requestYomitanScanTokens marks grouped entries when SubMiner dictionary alias only exists on definitions', async () => {
  const scripts: string[] = [];
  const deps = createScanDeps(
    (action, params) => {
      if (action === 'optionsGetFull') {
        return {
          profileCurrent: 0,
          profiles: [
            {
              options: {
                scanning: { length: 40 },
              },
            },
          ],
        };
      }
      if (action === 'getDictionaryInfo') {
        return [];
      }
      if (action === 'termsFind') {
        const text = (params as { text?: string } | undefined)?.text;
        if (text === 'カズマ') {
          return {
            originalTextLength: 3,
            dictionaryEntries: [
              {
                dictionaryAlias: '',
                headwords: [
                  {
                    term: 'カズマ',
                    reading: 'かずま',
                    sources: [{ originalText: 'カズマ', isPrimary: true, matchType: 'exact' }],
                  },
                ],
                definitions: [
                  { dictionary: 'JMdict', dictionaryAlias: 'JMdict' },
                  {
                    dictionary: 'SubMiner Character Dictionary (AniList 130298)',
                    dictionaryAlias: 'SubMiner Character Dictionary (AniList 130298)',
                  },
                ],
              },
            ],
          };
        }
        return { originalTextLength: 0, dictionaryEntries: [] };
      }
      throw new Error(`unexpected action: ${action}`);
    },
    { onScript: (script) => scripts.push(script) },
  );

  const result = await requestYomitanScanTokens(
    'カズマ',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  assert.ok(scripts.some((script) => script.includes('getPreferredHeadword')));
  assert.equal(Array.isArray(result), true);
  assert.equal((result as { length?: number } | null)?.length, 1);
  assert.equal((result as Array<{ surface?: string }>)[0]?.surface, 'カズマ');
  assert.equal((result as Array<{ headword?: string }>)[0]?.headword, 'カズマ');
  assert.equal((result as Array<{ startPos?: number }>)[0]?.startPos, 0);
  assert.equal((result as Array<{ endPos?: number }>)[0]?.endPos, 3);
  assert.equal((result as Array<{ isNameMatch?: boolean }>)[0]?.isNameMatch, true);
});

test('requestYomitanScanTokens ignores SubMiner character entries from other media', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }
    const text = (params as { text?: string } | undefined)?.text;
    if (text !== 'カズ') {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }
    return {
      originalTextLength: 2,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: 'カズ',
              reading: 'かず',
              sources: [{ originalText: 'カズ', isPrimary: true, matchType: 'exact' }],
            },
          ],
          definitions: [
            {
              dictionary: 'SubMiner Character Dictionary',
              dictionaryAlias: 'SubMiner Character Dictionary',
              entries: [
                {
                  type: 'structured-content',
                  content: {
                    tag: 'img',
                    path: 'img/m115230-c9.png',
                    alt: 'Kaz',
                  },
                },
              ],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens(
    'カズ',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true, currentCharacterDictionaryMediaId: 21202 },
  );

  // No dictionary-backed token survives (the only match belongs to another
  // media's character dictionary), so the line reports no tokenization.
  assert.equal(result, null);
});

test('requestYomitanScanTokens accepts SubMiner character entries with structured-content media data', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }
    const text = (params as { text?: string } | undefined)?.text;
    if (text !== 'アクア') {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }
    return {
      originalTextLength: 3,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: 'アクア',
              reading: 'あくあ',
              sources: [{ originalText: 'アクア', isPrimary: true, matchType: 'exact' }],
            },
          ],
          definitions: [
            {
              dictionary: 'SubMiner Character Dictionary',
              dictionaryAlias: 'SubMiner Character Dictionary',
              entries: [
                {
                  type: 'structured-content',
                  content: {
                    tag: 'div',
                    data: { subminerMediaId: '21699' },
                    content: [
                      {
                        tag: 'img',
                        path: 'img/m115230-c1.png',
                        alt: 'アクア',
                      },
                    ],
                  },
                },
              ],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens(
    'アクア',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true, currentCharacterDictionaryMediaId: 21699 },
  );

  assert.equal(Array.isArray(result), true);
  assert.equal((result as Array<{ surface?: string }>)[0]?.surface, 'アクア');
  assert.equal((result as Array<{ isNameMatch?: boolean }>)[0]?.isNameMatch, true);
});

test('requestYomitanScanTokens greedily tokenizes character names before longer generic matches', async () => {
  let scanCallScript = '';
  const nameEntry = (term: string, reading: string) => ({
    headwords: [
      {
        term,
        reading,
        sources: [{ originalText: term, isPrimary: true, matchType: 'exact' }],
      },
    ],
    definitions: [
      {
        dictionary: 'SubMiner Character Dictionary (AniList 130298)',
        dictionaryAlias: 'SubMiner Character Dictionary (AniList 130298)',
      },
    ],
  });
  const jmdictEntry = (term: string, reading: string, originalText: string) => ({
    headwords: [
      {
        term,
        reading,
        sources: [{ originalText, isPrimary: true, matchType: 'exact' }],
      },
    ],
    definitions: [{ dictionary: 'JMdict', dictionaryAlias: 'JMdict' }],
  });

  const deps = createScanDeps(
    (action, params) => {
      if (action === 'optionsGetFull') {
        return {
          profileCurrent: 0,
          profiles: [
            {
              options: {
                scanning: { length: 40 },
                dictionaries: [
                  { name: 'JMdict', enabled: true },
                  { name: 'SubMiner Character Dictionary (AniList 130298)', enabled: true },
                ],
              },
            },
          ],
        };
      }
      if (action === 'getDictionaryInfo') {
        return [];
      }
      if (action !== 'termsFind') {
        throw new Error(`unexpected action: ${action}`);
      }
      const text = (params as { text?: string } | undefined)?.text ?? '';
      if (text.startsWith('美姫')) {
        return { originalTextLength: 2, dictionaryEntries: [nameEntry('美姫', 'みき')] };
      }
      if (text.startsWith('とヨータ')) {
        // Greedy generic match: とヨー normalizes to とよう (渡洋). Without the
        // name pre-pass this consumes the ヨ of ヨータ.
        return {
          originalTextLength: 3,
          dictionaryEntries: [
            jmdictEntry('渡洋', 'とよう', 'とヨー'),
            jmdictEntry('と', 'と', 'と'),
          ],
        };
      }
      if (text.startsWith('ヨータ')) {
        return { originalTextLength: 3, dictionaryEntries: [nameEntry('ヨータ', 'よーた')] };
      }
      if (text === 'と') {
        return { originalTextLength: 1, dictionaryEntries: [jmdictEntry('と', 'と', 'と')] };
      }
      return { originalTextLength: 0, dictionaryEntries: [] };
    },
    {
      onScript: (script) => {
        if (script.includes('__subminerYomitanScan(')) {
          scanCallScript = script;
        }
      },
    },
  );

  const result = await requestYomitanScanTokens(
    '美姫とヨータ',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  assert.match(scanCallScript, /"greedyNameScanEnabled":true/);
  assert.equal(Array.isArray(result), true);
  assert.deepEqual(
    result?.map(({ surface, headword, startPos, endPos, isNameMatch }) => ({
      surface,
      headword,
      startPos,
      endPos,
      isNameMatch,
    })),
    [
      { surface: '美姫', headword: '美姫', startPos: 0, endPos: 2, isNameMatch: true },
      { surface: 'と', headword: 'と', startPos: 2, endPos: 3, isNameMatch: false },
      { surface: 'ヨータ', headword: 'ヨータ', startPos: 3, endPos: 6, isNameMatch: true },
    ],
  );
});

test('requestYomitanScanTokens lets a longer generic word beat a shorter name at the same position', async () => {
  const nameEntry = (term: string, reading: string) => ({
    headwords: [
      {
        term,
        reading,
        sources: [{ originalText: term, isPrimary: true, matchType: 'exact' }],
      },
    ],
    definitions: [
      {
        dictionary: 'SubMiner Character Dictionary (AniList 130298)',
        dictionaryAlias: 'SubMiner Character Dictionary (AniList 130298)',
      },
    ],
  });
  const jmdictEntry = (term: string, reading: string, originalText: string) => ({
    headwords: [
      {
        term,
        reading,
        sources: [{ originalText, isPrimary: true, matchType: 'exact' }],
      },
    ],
    definitions: [{ dictionary: 'JMdict', dictionaryAlias: 'JMdict' }],
  });

  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JMdict', enabled: true },
                { name: 'SubMiner Character Dictionary (AniList 130298)', enabled: true },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (text.startsWith('空気')) {
      // A character named 空 matches here, but the generic 空気 is longer and
      // must win the position.
      return {
        originalTextLength: 2,
        dictionaryEntries: [nameEntry('空', 'くう'), jmdictEntry('空気', 'くうき', '空気')],
      };
    }
    if (text.startsWith('変わって')) {
      return {
        originalTextLength: 4,
        dictionaryEntries: [jmdictEntry('変わる', 'かわる', '変わって')],
      };
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });

  const result = await requestYomitanScanTokens(
    '空気変わって',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  assert.equal(Array.isArray(result), true);
  assert.deepEqual(
    result?.map(({ surface, headword, startPos, endPos, isNameMatch }) => ({
      surface,
      headword,
      startPos,
      endPos,
      isNameMatch,
    })),
    [
      { surface: '空気', headword: '空気', startPos: 0, endPos: 2, isNameMatch: false },
      { surface: '変わって', headword: '変わる', startPos: 2, endPos: 6, isNameMatch: false },
    ],
  );
});

test('requestYomitanScanTokens lets a generic word beat a name it fully contains', async () => {
  const nameEntry = (term: string, reading: string) => ({
    headwords: [
      {
        term,
        reading,
        sources: [{ originalText: term, isPrimary: true, matchType: 'exact' }],
      },
    ],
    definitions: [
      {
        dictionary: 'SubMiner Character Dictionary (AniList 130298)',
        dictionaryAlias: 'SubMiner Character Dictionary (AniList 130298)',
      },
    ],
  });
  const jmdictEntry = (term: string, reading: string, originalText: string) => ({
    headwords: [
      {
        term,
        reading,
        sources: [{ originalText, isPrimary: true, matchType: 'exact' }],
      },
    ],
    definitions: [{ dictionary: 'JMdict', dictionaryAlias: 'JMdict' }],
  });

  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JMdict', enabled: true },
                { name: 'SubMiner Character Dictionary (AniList 130298)', enabled: true },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (text.startsWith('写真')) {
      return {
        originalTextLength: 2,
        dictionaryEntries: [jmdictEntry('写真', 'しゃしん', '写真')],
      };
    }
    if (text.startsWith('写')) {
      return { originalTextLength: 1, dictionaryEntries: [jmdictEntry('写', 'しゃ', '写')] };
    }
    if (text.startsWith('真')) {
      // The given name of 安田真 also matches the second half of 写真.
      return {
        originalTextLength: 1,
        dictionaryEntries: [nameEntry('真', 'しん'), jmdictEntry('真', 'しん', '真')],
      };
    }
    if (text.startsWith('は')) {
      return { originalTextLength: 1, dictionaryEntries: [jmdictEntry('は', 'は', 'は')] };
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });

  const result = await requestYomitanScanTokens(
    '写真は',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  assert.equal(Array.isArray(result), true);
  assert.deepEqual(
    result?.map(({ surface, headword, startPos, endPos, isNameMatch }) => ({
      surface,
      headword,
      startPos,
      endPos,
      isNameMatch,
    })),
    [
      { surface: '写真', headword: '写真', startPos: 0, endPos: 2, isNameMatch: false },
      { surface: 'は', headword: 'は', startPos: 2, endPos: 3, isNameMatch: false },
    ],
  );
});

test('requestYomitanScanTokens skips greedy name scan without an enabled character dictionary', async () => {
  let scanCallScript = '';
  const deps = createScanDeps(
    (action) => {
      if (action === 'optionsGetFull') {
        return {
          profileCurrent: 0,
          profiles: [
            {
              options: {
                scanning: { length: 40 },
                dictionaries: [{ name: 'JMdict', enabled: true }],
              },
            },
          ],
        };
      }
      if (action === 'getDictionaryInfo') {
        return [];
      }
      return { originalTextLength: 0, dictionaryEntries: [] };
    },
    {
      onScript: (script) => {
        if (script.includes('__subminerYomitanScan(')) {
          scanCallScript = script;
        }
      },
    },
  );

  await requestYomitanScanTokens(
    'アクア',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  assert.match(scanCallScript, /"greedyNameScanEnabled":false/);
});

test('requestYomitanScanTokens preserves matched headword word classes', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    const text = (params as { text?: string } | undefined)?.text;
    if (text !== 'は') {
      return { originalTextLength: 0, dictionaryEntries: [] };
    }

    return {
      originalTextLength: 1,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: 'は',
              reading: 'は',
              wordClasses: ['prt'],
              sources: [{ originalText: 'は', isPrimary: true, matchType: 'exact' }],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('は', deps, { error: () => undefined });

  assert.deepEqual((result as Array<{ wordClasses?: string[] }>)[0]?.wordClasses, ['prt']);
});

test('requestYomitanScanTokens skips fallback fragments without exact primary source matches', async () => {
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }

    {
      const text = (params as { text?: string } | undefined)?.text ?? '';
      if (text.startsWith('だが ')) {
        return {
          originalTextLength: 2,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term: 'だが',
                  reading: 'だが',
                  sources: [{ originalText: 'だが', isPrimary: true, matchType: 'exact' }],
                },
              ],
            },
          ],
        };
      }
      if (text.startsWith('それでも')) {
        return {
          originalTextLength: 4,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term: 'それでも',
                  reading: 'それでも',
                  sources: [{ originalText: 'それでも', isPrimary: true, matchType: 'exact' }],
                },
              ],
            },
          ],
        };
      }
      if (text.startsWith('届かぬ')) {
        return {
          originalTextLength: 3,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term: '届く',
                  reading: 'とどく',
                  sources: [{ originalText: '届かぬ', isPrimary: true, matchType: 'exact' }],
                },
              ],
            },
          ],
        };
      }
      if (text.startsWith('高み')) {
        return {
          originalTextLength: 2,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term: '高み',
                  reading: 'たかみ',
                  sources: [{ originalText: '高み', isPrimary: true, matchType: 'exact' }],
                },
              ],
            },
          ],
        };
      }
      if (text.startsWith('があった')) {
        return {
          originalTextLength: 2,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term: 'があ',
                  reading: '',
                  sources: [{ originalText: 'が', isPrimary: true, matchType: 'exact' }],
                },
              ],
            },
          ],
        };
      }
      if (text.startsWith('あった')) {
        return {
          originalTextLength: 3,
          dictionaryEntries: [
            {
              headwords: [
                {
                  term: 'ある',
                  reading: 'ある',
                  sources: [{ originalText: 'あった', isPrimary: true, matchType: 'exact' }],
                },
              ],
            },
          ],
        };
      }
      return { originalTextLength: 0, dictionaryEntries: [] };
    }
  });

  const result = await requestYomitanScanTokens('だが それでも届かぬ高みがあった', deps, {
    error: () => undefined,
  });

  assert.deepEqual(
    result?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      startPos: token.startPos,
      endPos: token.endPos,
    })),
    [
      {
        surface: 'だが',
        headword: 'だが',
        startPos: 0,
        endPos: 2,
      },
      {
        surface: 'それでも',
        headword: 'それでも',
        startPos: 3,
        endPos: 7,
      },
      {
        surface: '届かぬ',
        headword: '届く',
        startPos: 7,
        endPos: 10,
      },
      {
        surface: '高み',
        headword: '高み',
        startPos: 10,
        endPos: 12,
      },
      // が has no exact primary source match, so it survives only as an
      // unparsed filler run (the parseText segmentation used to supply this).
      {
        surface: 'が',
        headword: 'が',
        startPos: 12,
        endPos: 13,
      },
      {
        surface: 'あった',
        headword: 'ある',
        startPos: 13,
        endPos: 16,
      },
    ],
  );
  assert.equal(result?.[4]?.isUnparsedRun, true);
});

function createSingleTermScanHandler(lookups: string[]) {
  return (action: string, params: unknown): unknown => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    if (action !== 'termsFind') {
      throw new Error(`unexpected action: ${action}`);
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    lookups.push(text);
    if (text.startsWith('猫')) {
      return {
        originalTextLength: 1,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: '猫',
                reading: 'ねこ',
                sources: [{ originalText: '猫', isPrimary: true, matchType: 'exact' }],
              },
            ],
          },
        ],
      };
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  };
}

test('requestYomitanScanTokens reuses the cross-line termsFind cache for repeated lookups', async () => {
  const lookups: string[] = [];
  const deps = createScanDeps(createSingleTermScanHandler(lookups));

  const first = await requestYomitanScanTokens('猫', deps, { error: () => undefined });
  const second = await requestYomitanScanTokens('猫', deps, { error: () => undefined });

  assert.equal(first?.length, 1);
  assert.equal(second?.length, 1);
  // The second line hits the window-persistent cache: no new backend lookup.
  assert.equal(countTermsFindLookups(lookups, '猫'), 1);
});

test('clearYomitanParserCachesForWindow invalidates the cross-line termsFind cache', async () => {
  const lookups: string[] = [];
  const deps = createScanDeps(createSingleTermScanHandler(lookups));

  await requestYomitanScanTokens('猫', deps, { error: () => undefined });
  clearYomitanParserCachesForWindow(deps.getYomitanParserWindow() as never);
  await requestYomitanScanTokens('猫', deps, { error: () => undefined });

  assert.equal(countTermsFindLookups(lookups, '猫'), 2);
});

test('an oversized termsFind result is dropped from the cache instead of being reused', async () => {
  const lookups: string[] = [];
  // One entry over the runtime's 20,000 retained-entry budget: the weight is
  // only known once the lookup resolves, so the cache has to re-check then.
  const oversizedEntries = Array.from({ length: 20_001 }, () => ({
    headwords: [
      {
        term: '猫',
        reading: 'ねこ',
        sources: [{ originalText: '猫', isPrimary: true, matchType: 'exact' }],
      },
    ],
  }));
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    lookups.push(text);
    if (text.startsWith('猫')) {
      return { originalTextLength: 1, dictionaryEntries: oversizedEntries };
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });

  await requestYomitanScanTokens('猫', deps, { error: () => undefined });
  await requestYomitanScanTokens('猫', deps, { error: () => undefined });

  assert.equal(countTermsFindLookups(lookups, '猫'), 2);
});

test('scanner tokens survive a retry-budget escalation whose parseText finds nothing', async () => {
  const parsedTexts: string[] = [];
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (action === 'parseText') {
      parsedTexts.push(text);
      return [];
    }
    if (text.startsWith('猫')) {
      return {
        originalTextLength: 1,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: '猫',
                reading: 'ねこ',
                sources: [{ originalText: '猫', isPrimary: true, matchType: 'exact' }],
              },
            ],
          },
        ],
      };
    }
    // The rest of the line burns the blind-retry budget at every position.
    return {
      originalTextLength: text.length,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: 'ミスマッチ',
              reading: 'みすまっち',
              sources: [{ originalText: 'ZZZ', isPrimary: true, matchType: 'exact' }],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('猫あいうえおかきくけこ', deps, {
    error: () => undefined,
  });

  // The escalation ran exactly once and found nothing, so the tokens the
  // scanner did resolve are kept instead of dropping the line to raw text.
  assert.deepEqual(parsedTexts, ['猫あいうえおかきくけこ']);
  assert.equal(result?.[0]?.surface, '猫');
});

test('requestYomitanScanTokens skips termsFind lookups at punctuation and whitespace positions', async () => {
  const lookups: string[] = [];
  const deps = createScanDeps(createSingleTermScanHandler(lookups));

  const result = await requestYomitanScanTokens('「猫」…♪', deps, { error: () => undefined });

  assert.equal(result?.length, 1);
  assert.equal(result?.[0]?.surface, '猫');
  assert.equal(countTermsFindLookups(lookups, '猫'), 1);
  for (const skipped of ['「', '」', '…', '♪']) {
    assert.equal(countTermsFindLookups(lookups, skipped), 0, `expected no lookup at ${skipped}`);
  }
});

test('requestYomitanScanTokens caps blind retries and escalates the line to parseText', async () => {
  const lookups: string[] = [];
  const parsedTexts: string[] = [];
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (action === 'parseText') {
      parsedTexts.push(text);
      return [
        {
          source: 'scanning-parser',
          index: 0,
          content: [
            [{ text: 'あいうえお', reading: 'あいうえお', headwords: [[{ term: 'あい' }]] }],
          ],
        },
      ];
    }
    lookups.push(text);
    // Every window "matches" its whole length but never yields an
    // exact-source headword, the worst case for the retry ladder: each step
    // down is a blind guess with nothing shorter reported to aim at.
    return {
      originalTextLength: text.length,
      dictionaryEntries: [
        {
          headwords: [
            {
              term: 'ミスマッチ',
              reading: 'みすまっち',
              sources: [{ originalText: 'ZZZ', isPrimary: true, matchType: 'exact' }],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('あいうえおかきくけこ', deps, {
    error: () => undefined,
  });

  // Position 0: one initial window lookup plus at most four blind retries, so
  // the ladder cannot degrade into a lookup per window length.
  assert.equal(countTermsFindLookups(lookups, 'あいうえお'), 5);
  // Giving up there would leave the line unparsed, so it escalates to the one
  // full parse the scanner normally replaces.
  assert.deepEqual(parsedTexts, ['あいうえおかきくけこ']);
  assert.equal(result?.[0]?.headword, 'あい');
});

test('requestYomitanScanTokens keeps shrinking while the backend guides the retry ladder', async () => {
  const lookups: string[] = [];
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    lookups.push(text);
    // Normalization keeps eating one character past the term, so every window
    // reports a shorter consumed length: informative steps that must not be
    // spent from the blind-retry budget. The term only surfaces at length 2,
    // six lookups down the ladder.
    if (text.length === 2) {
      return {
        originalTextLength: 2,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: 'あい',
                reading: 'あい',
                sources: [{ originalText: 'あい', isPrimary: true, matchType: 'exact' }],
              },
            ],
          },
        ],
      };
    }
    return {
      originalTextLength: Math.max(text.length - 1, 0),
      dictionaryEntries: [
        {
          headwords: [
            {
              term: 'ミスマッチ',
              reading: 'みすまっち',
              sources: [{ originalText: 'ZZZ', isPrimary: true, matchType: 'exact' }],
            },
          ],
        },
      ],
    };
  });

  const result = await requestYomitanScanTokens('あいうえおかきくけこさしすせ', deps, {
    error: () => undefined,
  });

  assert.equal(result?.[0]?.surface, 'あい');
  // Windows of 14, 12, 10, 8, 6, 4 characters, then the match at 2: a ladder
  // capped at four lookups would stop at 6 and leave the line unparsed.
  assert.equal(countTermsFindLookups(lookups, 'あい'), 7);
});

test('requestYomitanScanTokens falls back to parseText when the scanner eval fails', async () => {
  const deps = createDeps(async (script) => {
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profiles: [{ options: { scanning: { length: 40 } } }],
      };
    }
    if (script.includes('__subminerYomitanScan(')) {
      throw new Error('eval failed');
    }
    if (script.includes('parseText')) {
      return [
        {
          source: 'scanning-parser',
          index: 0,
          content: [
            [
              {
                text: '取り組んで',
                reading: 'とりくんで',
                headwords: [[{ term: '取り組む' }]],
              },
            ],
          ],
        },
      ];
    }
    return null;
  });

  const errors: string[] = [];
  const result = await requestYomitanScanTokens('取り組んで', deps, {
    error: (message) => errors.push(message),
  });

  assert.deepEqual(result, [
    {
      surface: '取り組んで',
      reading: 'とりくんで',
      headword: '取り組む',
      startPos: 0,
      endPos: 5,
    },
  ]);
  assert.equal(errors.length, 1);
});

test('getYomitanDictionaryInfo requests dictionary info via backend action', async () => {
  let scriptValue = '';
  const deps = createDeps(async (script) => {
    scriptValue = script;
    return [{ title: 'SubMiner Character Dictionary (AniList 130298)', revision: '1' }];
  });

  const dictionaries = await getYomitanDictionaryInfo(deps, { error: () => undefined });
  assert.equal(dictionaries.length, 1);
  assert.equal(dictionaries[0]?.title, 'SubMiner Character Dictionary (AniList 130298)');
  assert.match(scriptValue, /getDictionaryInfo/);
});

test('dictionary settings helpers upsert and remove dictionary entries without reordering', async () => {
  const scripts: string[] = [];
  const optionsFull = {
    profileCurrent: 0,
    profiles: [
      {
        options: {
          dictionaries: [
            {
              name: 'Jitendex',
              alias: 'Jitendex',
              enabled: true,
            },
            {
              name: 'SubMiner Character Dictionary (AniList 1)',
              alias: 'SubMiner Character Dictionary (AniList 1)',
              enabled: false,
            },
          ],
        },
      },
    ],
  };

  const deps = createDeps(async (script) => {
    scripts.push(script);
    if (script.includes('optionsGetFull')) {
      return structuredClone(optionsFull);
    }
    if (script.includes('setAllSettings')) {
      return true;
    }
    return null;
  });

  const title = 'SubMiner Character Dictionary (AniList 1)';
  const upserted = await upsertYomitanDictionarySettings(title, 'all', deps, {
    error: () => undefined,
  });
  const removed = await removeYomitanDictionarySettings(title, 'all', 'delete', deps, {
    error: () => undefined,
  });

  assert.equal(upserted, true);
  assert.equal(removed, true);
  const setCalls = scripts.filter((script) => script.includes('setAllSettings')).length;
  assert.equal(setCalls, 2);

  const upsertScript = scripts.find(
    (script) =>
      script.includes('setAllSettings') &&
      script.includes('"SubMiner Character Dictionary (AniList 1)"'),
  );
  assert.ok(upsertScript);
  const jitendexOffset = upsertScript?.indexOf('"Jitendex"') ?? -1;
  const subMinerOffset = upsertScript?.indexOf('"SubMiner Character Dictionary (AniList 1)"') ?? -1;
  assert.equal(jitendexOffset >= 0, true);
  assert.equal(subMinerOffset >= 0, true);
  assert.equal(jitendexOffset < subMinerOffset, true);
  assert.match(upsertScript ?? '', /"enabled":true/);
});

test('importYomitanDictionaryFromZip imports via localhost URL instead of embedding archive bytes in script', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yomitan-import-'));
  const zipPath = path.join(tempDir, 'dict.zip');
  fs.writeFileSync(zipPath, Buffer.from('zip-bytes'));

  const scripts: string[] = [];
  const servedArchives: string[] = [];
  const settingsWindow = {
    isDestroyed: () => false,
    destroy: () => undefined,
    webContents: {
      executeJavaScript: async (script: string) => {
        scripts.push(script);
        const urlMatch = script.match(/importDictionaryArchiveUrl\(\s*"([^"]+)"/);
        if (urlMatch) {
          const response = await fetch(JSON.parse(`"${urlMatch[1]}"`) as string);
          servedArchives.push(await response.text());
        }
        return true;
      },
    },
  };

  const deps = createDeps(async () => true, {
    createYomitanExtensionWindow: async (pageName: string) => {
      assert.equal(pageName, 'settings.html');
      return settingsWindow;
    },
  });

  const imported = await importYomitanDictionaryFromZip(zipPath, deps, {
    error: () => undefined,
  });

  assert.equal(imported, true);
  assert.equal(
    scripts.some((script) => script.includes('__subminerYomitanSettingsAutomation')),
    true,
  );
  assert.equal(
    scripts.some((script) => script.includes('importDictionaryArchiveUrl')),
    true,
  );
  assert.deepEqual(servedArchives, ['zip-bytes']);
  assert.equal(
    scripts.some((script) => script.includes('emlwLWJ5dGVz')),
    false,
  );
  assert.equal(
    scripts.some((script) => script.includes('subminerImportDictionary')),
    false,
  );
});

test('importYomitanDictionaryFromZip falls back to base64 import for older Yomitan bridge', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yomitan-import-'));
  const zipPath = path.join(tempDir, 'dict.zip');
  fs.writeFileSync(zipPath, Buffer.from('zip-bytes'));

  const scripts: string[] = [];
  const settingsWindow = {
    isDestroyed: () => false,
    destroy: () => undefined,
    webContents: {
      executeJavaScript: async (script: string) => {
        scripts.push(script);
        if (
          script.includes(
            'typeof globalThis.__subminerYomitanSettingsAutomation.importDictionaryArchiveUrl',
          )
        ) {
          return false;
        }
        return true;
      },
    },
  };

  const deps = createDeps(async () => true, {
    createYomitanExtensionWindow: async (pageName: string) => {
      assert.equal(pageName, 'settings.html');
      return settingsWindow;
    },
  });

  const imported = await importYomitanDictionaryFromZip(zipPath, deps, {
    error: () => undefined,
  });

  assert.equal(imported, true);
  assert.equal(
    scripts.some((script) => script.includes('importDictionaryArchiveBase64')),
    true,
  );
  assert.equal(
    scripts.some((script) => script.includes('importDictionaryArchiveUrl(')),
    false,
  );
  assert.equal(
    scripts.some((script) => script.includes('emlwLWJ5dGVz')),
    true,
  );
});

test('importYomitanDictionaryFromZip returns false when served archive cannot be read', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yomitan-import-'));
  const zipPath = path.join(tempDir, 'dict.zip');
  fs.writeFileSync(zipPath, Buffer.from('zip-bytes'));

  const settingsWindow = {
    isDestroyed: () => false,
    destroy: () => undefined,
    webContents: {
      executeJavaScript: async (script: string) => {
        const urlMatch = script.match(/importDictionaryArchiveUrl\(\s*"([^"]+)"/);
        if (!urlMatch) return true;
        fs.unlinkSync(zipPath);
        const response = await fetch(JSON.parse(`"${urlMatch[1]}"`) as string);
        return response.ok;
      },
    },
  };

  const deps = createDeps(async () => true, {
    createYomitanExtensionWindow: async (pageName: string) => {
      assert.equal(pageName, 'settings.html');
      return settingsWindow;
    },
  });

  const imported = await importYomitanDictionaryFromZip(zipPath, deps, {
    error: () => undefined,
  });

  assert.equal(imported, false);
});

test('deleteYomitanDictionaryByTitle uses settings automation bridge instead of custom backend action', async () => {
  const scripts: string[] = [];
  const settingsWindow = {
    isDestroyed: () => false,
    destroy: () => undefined,
    webContents: {
      executeJavaScript: async (script: string) => {
        scripts.push(script);
        return true;
      },
    },
  };

  const deps = createDeps(async () => true, {
    createYomitanExtensionWindow: async (pageName: string) => {
      assert.equal(pageName, 'settings.html');
      return settingsWindow;
    },
  });

  const deleted = await deleteYomitanDictionaryByTitle(
    'SubMiner Character Dictionary (AniList 130298)',
    deps,
    { error: () => undefined },
  );

  assert.equal(deleted, true);
  assert.equal(
    scripts.some((script) => script.includes('__subminerYomitanSettingsAutomation')),
    true,
  );
  assert.equal(
    scripts.some((script) => script.includes('deleteDictionary')),
    true,
  );
  assert.equal(
    scripts.some((script) => script.includes('subminerDeleteDictionary')),
    false,
  );
});

test('addYomitanNoteViaSearch returns note and duplicate ids from the bridge payload', async () => {
  const deps = createDeps(async (_script) => ({
    noteId: 42,
    duplicateNoteIds: [18, 7, 18],
  }));

  const result = await addYomitanNoteViaSearch('食べる', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, {
    noteId: 42,
    duplicateNoteIds: [18, 7, 18],
  });
});

test('addYomitanNoteViaSearch rejects invalid numeric note ids from the bridge shortcut', async () => {
  const deps = createDeps(async () => NaN);

  const result = await addYomitanNoteViaSearch('食べる', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, {
    noteId: null,
    duplicateNoteIds: [],
  });
});

test('addYomitanNoteViaSearch sanitizes invalid payload note ids while keeping valid duplicate ids', async () => {
  const deps = createDeps(async (_script) => ({
    noteId: -1,
    duplicateNoteIds: [18, 0, 7.5, 7],
  }));

  const result = await addYomitanNoteViaSearch('食べる', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, {
    noteId: null,
    duplicateNoteIds: [18, 7],
  });
});

test('requestYomitanScanTokens still finds an emphatically elongated name a longer generic match would swallow', async () => {
  // Yomitan collapses emphatic sequences, so ミナァァト resolves to the ミナト
  // entry. The generic word とミナ starts earlier and would swallow the name
  // unless the pre-pass reserves it, so this only passes when the candidate
  // prefilter still treats the elongated spelling as a possible name start.
  const deps = createScanDeps((action, params) => {
    if (action === 'optionsGetFull') {
      return {
        profileCurrent: 0,
        profiles: [
          {
            options: {
              scanning: { length: 40 },
              dictionaries: [
                { name: 'JMdict', enabled: true, id: 0 },
                { name: 'SubMiner Character Dictionary (AniList 1)', enabled: true, id: 1 },
              ],
            },
          },
        ],
      };
    }
    if (action === 'getDictionaryInfo') {
      return [];
    }
    const text = (params as { text?: string } | undefined)?.text ?? '';
    if (text.startsWith('とミナ')) {
      return {
        originalTextLength: 3,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: 'トミナ',
                reading: 'とみな',
                sources: [{ originalText: 'とミナ', isPrimary: true, matchType: 'exact' }],
              },
            ],
            definitions: [{ dictionary: 'JMdict' }],
          },
        ],
      };
    }
    if (text.startsWith('ミナァァト')) {
      return {
        originalTextLength: 5,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: 'ミナト',
                reading: 'みなと',
                sources: [{ originalText: 'ミナァァト', isPrimary: true, matchType: 'exact' }],
              },
            ],
            definitions: [{ dictionary: 'SubMiner Character Dictionary (AniList 1)' }],
          },
        ],
      };
    }
    if (text.startsWith('と')) {
      return {
        originalTextLength: 1,
        dictionaryEntries: [
          {
            headwords: [
              {
                term: 'と',
                reading: 'と',
                sources: [{ originalText: 'と', isPrimary: true, matchType: 'exact' }],
              },
            ],
            definitions: [{ dictionary: 'JMdict' }],
          },
        ],
      };
    }
    return { originalTextLength: 0, dictionaryEntries: [] };
  });

  const result = await requestYomitanScanTokens(
    'とミナァァト',
    deps,
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      nameCandidates: { key: 'media-1', forms: ['ミナト', 'みなと'] },
    },
  );

  const nameToken = result?.find((token) => token.isNameMatch === true);
  assert.ok(nameToken, 'expected the elongated name to be reserved by the pre-pass');
  assert.equal(nameToken?.headword, 'ミナト');
  assert.equal(nameToken?.startPos, 1);
});
