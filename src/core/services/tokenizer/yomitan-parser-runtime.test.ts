import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import * as vm from 'node:vm';
import {
  countTermsFindLookups,
  createBackendDeps,
  createDeps,
  createSettingsAutomationDeps,
  termEntry,
  termsFound,
} from './yomitan-scan-test-harness';
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

const TARGET_ANKI_SERVER = 'http://127.0.0.1:8766';
const CUSTOM_ANKI_SERVER = 'http://192.168.1.5:8765';
const CHARACTER_DICTIONARY = 'SubMiner Character Dictionary (AniList 130298)';

test('Yomitan restores direct Anki after disabling its managed proxy without a page helper', async () => {
  const options = { profiles: [{ options: { anki: { server: 'http://127.0.0.1:8765' } } }] };
  let managedUrl: string | null = null;
  const context = vm.createContext({
    chrome: {
      storage: {
        local: {
          get: async () => ({ subminerAnkiProxyUrl: managedUrl }),
          set: async (value: { subminerAnkiProxyUrl: string | null }) => {
            managedUrl = value.subminerAnkiProxyUrl;
          },
        },
      },
      runtime: {
        sendMessage: (
          message: { action: string },
          callback: (response: { result: unknown }) => void,
        ) => callback({ result: message.action === 'optionsGetFull' ? options : null }),
      },
    },
  });
  const deps = createDeps(async (script) =>
    structuredClone(await vm.runInContext(script, context)),
  );
  const logger = { error: assert.fail };
  assert.equal(
    await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', deps, logger, {
      forceOverride: true,
    }),
    true,
  );
  assert.equal(managedUrl, 'http://127.0.0.1:8766');
  assert.equal(await syncYomitanDefaultAnkiServer('http://127.0.0.1:8765', deps, logger), true);
  assert.equal(options.profiles[0]?.options.anki.server, 'http://127.0.0.1:8765');
  assert.equal(managedUrl, null);
});

for (const failure of ['optionsGetFull', 'setAllSettings', 'storageSet']) {
  test(`Yomitan retries disabling its managed proxy after ${failure} fails`, async () => {
    const proxyUrl = 'http://127.0.0.1:8766';
    const directUrl = 'http://127.0.0.1:8765';
    let options = { profiles: [{ options: { anki: { server: proxyUrl } } }] };
    let managedUrl: string | null = proxyUrl;
    let shouldFail = true;
    const context = vm.createContext({
      chrome: {
        storage: {
          local: {
            get: async () => ({ subminerAnkiProxyUrl: managedUrl }),
            set: async (value: { subminerAnkiProxyUrl: string | null }) => {
              if (shouldFail && failure === 'storageSet') throw new Error('Storage unavailable');
              managedUrl = value.subminerAnkiProxyUrl;
            },
          },
        },
        runtime: {
          sendMessage: (
            message: { action: string; params?: { value: typeof options } },
            callback: (response: { result?: unknown; error?: { message: string } }) => void,
          ) => {
            if (shouldFail && message.action === failure) {
              callback({ error: { message: 'Settings unavailable' } });
            } else if (message.action === 'optionsGetFull') {
              callback({ result: structuredClone(options) });
            } else if (message.action === 'setAllSettings' && message.params) {
              options = structuredClone(message.params.value);
              callback({ result: null });
            } else {
              assert.fail(`Unexpected action: ${message.action}`);
            }
          },
        },
      },
    });
    const deps = createDeps(async (script) =>
      structuredClone(await vm.runInContext(script, context)),
    );
    const errors: string[] = [];
    const logger = { error: (message: string) => errors.push(message) };

    assert.equal(await syncYomitanDefaultAnkiServer(directUrl, deps, logger), false);
    assert.equal(managedUrl, proxyUrl);
    assert.equal(
      options.profiles[0]?.options.anki.server,
      failure === 'storageSet' ? directUrl : proxyUrl,
    );
    assert.equal(errors.length, 1);

    shouldFail = false;
    assert.equal(await syncYomitanDefaultAnkiServer(directUrl, deps, logger), true);
    assert.equal(options.profiles[0]?.options.anki.server, directUrl);
    assert.equal(managedUrl, null);
    assert.equal(errors.length, 1);
  });
}

interface AnkiSyncProfile {
  profileCurrent: number;
  profiles: Array<{ options: { anki: Record<string, unknown> } }>;
}

// Runs the real sync script against a backend whose settings start as
// `optionsFull`; every setAllSettings payload lands in `saved`.
function createAnkiSyncDeps(optionsFull: AnkiSyncProfile) {
  const saved: AnkiSyncProfile[] = [];
  const deps = createBackendDeps({
    actions: {
      optionsGetFull: () => structuredClone(optionsFull),
      setAllSettings: (params) => {
        saved.push((params as { value: AnkiSyncProfile }).value);
        return true;
      },
    },
  });
  return { deps, saved };
}

function ankiProfile(anki: Record<string, unknown>): AnkiSyncProfile {
  return { profileCurrent: 0, profiles: [{ options: { anki } }] };
}

const ANKI_SERVER_SYNC_CASES = [
  {
    name: 'replaces the stock default server',
    currentServer: 'http://127.0.0.1:8765',
    forceOverride: false,
    expected: true,
    savedServers: [TARGET_ANKI_SERVER],
    infoLog: /Updated Yomitan default profile Anki server/,
  },
  {
    name: 'reports success without saving when already on the target server',
    currentServer: TARGET_ANKI_SERVER,
    forceOverride: false,
    expected: true,
    savedServers: [],
    infoLog: null,
  },
  {
    name: 'refuses to replace a custom server',
    currentServer: CUSTOM_ANKI_SERVER,
    forceOverride: false,
    expected: false,
    savedServers: [],
    infoLog: /blocked-existing-server/,
  },
  {
    name: 'replaces a custom server when forced',
    currentServer: CUSTOM_ANKI_SERVER,
    forceOverride: true,
    expected: true,
    savedServers: [TARGET_ANKI_SERVER],
    infoLog: /Updated Yomitan default profile Anki server/,
  },
];

for (const c of ANKI_SERVER_SYNC_CASES) {
  test(`syncYomitanDefaultAnkiServer ${c.name}`, async () => {
    const { deps, saved } = createAnkiSyncDeps(ankiProfile({ server: c.currentServer }));
    const infoLogs: string[] = [];

    const synced = await syncYomitanDefaultAnkiServer(
      TARGET_ANKI_SERVER,
      deps,
      { error: () => undefined, info: (message) => infoLogs.push(message) },
      { forceOverride: c.forceOverride },
    );

    assert.equal(synced, c.expected);
    assert.deepEqual(
      saved.map((value) => value.profiles[0]?.options.anki.server),
      c.savedServers,
    );
    if (c.infoLog) {
      assert.equal(infoLogs.length, 1);
      assert.match(infoLogs[0] ?? '', c.infoLog);
    } else {
      assert.deepEqual(infoLogs, []);
    }
  });
}

test('syncYomitanDefaultAnkiServer updates the active profile Anki deck', async () => {
  const { deps, saved } = createAnkiSyncDeps(
    ankiProfile({
      server: TARGET_ANKI_SERVER,
      cardFormats: [
        { type: 'term', deck: 'Default', model: 'Mining Note', fields: {} },
        { type: 'kanji', deck: 'Kanji', model: 'Kanji Note', fields: {} },
      ],
      terms: { deck: 'Default', model: 'Legacy Note', fields: {} },
    }),
  );

  const synced = await syncYomitanDefaultAnkiServer(
    TARGET_ANKI_SERVER,
    deps,
    { error: () => undefined, info: () => undefined },
    { deck: 'Minecraft', forceOverride: true },
  );

  assert.equal(synced, true);
  const anki = saved[0]?.profiles[0]?.options.anki as {
    cardFormats: Array<{ deck: string }>;
    terms: { deck: string };
  };
  assert.deepEqual(
    anki.cardFormats.map((format) => format.deck),
    ['Minecraft', 'Kanji'],
  );
  assert.equal(anki.terms.deck, 'Minecraft');
});

test('syncYomitanDefaultAnkiServer logs and returns false on script failure', async () => {
  const deps = createDeps(async () => {
    throw new Error('execute failed');
  });

  const errorLogs: string[] = [];
  const updated = await syncYomitanDefaultAnkiServer(TARGET_ANKI_SERVER, deps, {
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

type TermReadingPair = { term: string; reading: string | null };

// Frequency backend over one enabled dictionary. `lookup` answers each
// getTermFrequencies request; the requested pair lists land in `requests`.
function createFrequencyDeps(
  lookup: (termReadingList: TermReadingPair[]) => unknown[],
  options: { dictionary?: string; dictionaryInfo?: unknown[] } = {},
) {
  const requests: TermReadingPair[][] = [];
  const actionLog: string[] = [];
  const deps = createBackendDeps({
    dictionaries: [options.dictionary ?? 'freq-dict'],
    dictionaryInfo: options.dictionaryInfo,
    actionLog,
    actions: {
      getTermFrequencies: (params) => {
        const { termReadingList } = params as { termReadingList: TermReadingPair[] };
        requests.push(termReadingList);
        return lookup(termReadingList);
      },
    },
  });
  return { deps, requests, actionLog };
}

function frequencyEntry(
  term: string,
  reading: string | null,
  frequency: number,
  displayValue: unknown = String(frequency),
  extra: Record<string, unknown> = {},
) {
  return { term, reading, dictionary: 'freq-dict', frequency, displayValue, ...extra };
}

const quietLogger = { error: () => undefined };

test('requestYomitanTermFrequencies returns normalized frequency entries', async () => {
  const { deps } = createFrequencyDeps(() => [
    frequencyEntry('猫', 'ねこ', 77, '77', { hasReading: true }),
    frequencyEntry('鍛える', 'きたえる', 46961, '2847,46961', { hasReading: false }),
    { term: 'invalid', dictionary: 'freq-dict', frequency: 0 },
  ]);

  const result = await requestYomitanTermFrequencies(
    [{ term: '猫', reading: 'ねこ' }],
    deps,
    quietLogger,
  );

  assert.deepEqual(
    result.map(({ term, hasReading, frequency, dictionaryPriority }) => ({
      term,
      hasReading,
      frequency,
      dictionaryPriority,
    })),
    [
      { term: '猫', hasReading: true, frequency: 77, dictionaryPriority: 0 },
      { term: '鍛える', hasReading: false, frequency: 2847, dictionaryPriority: 0 },
    ],
  );
});

const DISPLAY_VALUE_RANK_CASES = [
  { name: 'array pair', frequency: 157632, displayValue: [7141, 157632], expected: 7141 },
  { name: 'string with leading digits', frequency: 1234, displayValue: '1,234', expected: 1 },
];

for (const c of DISPLAY_VALUE_RANK_CASES) {
  test(`requestYomitanTermFrequencies takes the primary rank from a displayValue ${c.name}`, async () => {
    const { deps } = createFrequencyDeps(() => [
      frequencyEntry('例', 'れい', c.frequency, c.displayValue),
    ]);

    const result = await requestYomitanTermFrequencies(
      [{ term: '例', reading: 'れい' }],
      deps,
      quietLogger,
    );

    assert.deepEqual(
      result.map(({ term, frequency }) => ({ term, frequency })),
      [{ term: '例', frequency: c.expected }],
    );
  });
}

test('requestYomitanTermFrequencies ignores occurrence-based dictionaries for rank tagging', async () => {
  const { deps } = createFrequencyDeps(
    () => [{ ...frequencyEntry('潜む', 'ひそむ', 118121, null), dictionary: 'CC100' }],
    {
      dictionary: 'CC100',
      dictionaryInfo: [{ title: 'CC100', frequencyMode: 'occurrence-based' }],
    },
  );

  const result = await requestYomitanTermFrequencies(
    [{ term: '潜む', reading: 'ひそむ' }],
    deps,
    quietLogger,
  );

  assert.deepEqual(result, []);
});

test('requestYomitanTermFrequencies requests term-only fallback only after reading miss', async () => {
  const { deps, requests } = createFrequencyDeps(([pair]) =>
    pair?.reading === null ? [frequencyEntry('断じて', null, 7082)] : [],
  );

  const result = await requestYomitanTermFrequencies(
    [{ term: '断じて', reading: 'だん' }],
    deps,
    quietLogger,
  );

  assert.deepEqual(
    result.map((entry) => entry.frequency),
    [7082],
  );
  assert.deepEqual(requests, [
    [{ term: '断じて', reading: 'だん' }],
    [{ term: '断じて', reading: null }],
  ]);
});

test('requestYomitanTermFrequencies avoids term-only fallback request when reading lookup succeeds', async () => {
  const { deps, requests } = createFrequencyDeps(() => [
    frequencyEntry('鍛える', 'きたえる', 2847),
  ]);

  const result = await requestYomitanTermFrequencies(
    [{ term: '鍛える', reading: 'きた' }],
    deps,
    quietLogger,
  );

  assert.equal(result.length, 1);
  assert.deepEqual(requests, [[{ term: '鍛える', reading: 'きた' }]]);
});

test('requestYomitanTermFrequencies caches profile metadata between calls', async () => {
  const { deps, actionLog } = createFrequencyDeps(([pair]) => [
    frequencyEntry(pair?.term ?? '', pair?.reading ?? null, 12),
  ]);

  await requestYomitanTermFrequencies([{ term: '猫', reading: 'ねこ' }], deps, quietLogger);
  await requestYomitanTermFrequencies([{ term: '犬', reading: 'いぬ' }], deps, quietLogger);

  assert.equal(actionLog.filter((action) => action === 'optionsGetFull').length, 1);
});

test('requestYomitanTermFrequencies caches repeated term+reading lookups', async () => {
  const { deps, requests } = createFrequencyDeps(() => [frequencyEntry('猫', 'ねこ', 77)]);

  await requestYomitanTermFrequencies([{ term: '猫', reading: 'ねこ' }], deps, quietLogger);
  await requestYomitanTermFrequencies([{ term: '猫', reading: 'ねこ' }], deps, quietLogger);

  assert.equal(requests.length, 1);
});

test('requestYomitanScanTokens tokenizes with the in-window scanner and no parseText request', async () => {
  const parsedTexts: string[] = [];
  const deps = createBackendDeps({
    termsFind: (text) =>
      text.startsWith('取り組んで')
        ? termsFound(5, termEntry('取り組む', 'とりくむ', { originalText: '取り組んで' }))
        : null,
    actions: {
      parseText: (params) => {
        parsedTexts.push((params as { text: string }).text);
        return [];
      },
    },
  });

  const result = await requestYomitanScanTokens('取り組んで', deps, quietLogger);

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
  // The scanner walk is the only tokenization request: no duplicate full parse.
  assert.deepEqual(parsedTexts, []);
});

test('requestYomitanScanTokens warns when active Yomitan profile has no dictionaries', async () => {
  const warnings: Array<{ message: string; details: unknown }> = [];
  const deps = createBackendDeps({ dictionaries: [], termsFind: () => null });

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
  // 待ち合わせ matches, the trailing る does not, so the kana run extends the
  // previous token instead of becoming its own filler token.
  const deps = createBackendDeps({
    termsFind: (text) =>
      text.startsWith('待ち合わせ')
        ? termsFound(5, termEntry('待ち合わせる', 'まちあわせる', { originalText: '待ち合わせ' }))
        : null,
  });

  const result = await requestYomitanScanTokens('待ち合わせる', deps, quietLogger);

  assert.equal(result?.length, 1);
  assert.equal(result?.[0]?.surface, '待ち合わせる');
  assert.equal(result?.[0]?.endPos, 6);
  // The reading must grow with the surface: a short reading fails
  // isCompleteReadingForSurface and silently disables the known-word reading
  // fallback downstream.
  assert.equal(result?.[0]?.reading, 'まちあわせる');
});

test('requestYomitanScanTokens emits unparsed filler runs for text the scanner skips', async () => {
  const deps = createBackendDeps({
    termsFind: (text) => {
      if (text.startsWith('や')) {
        return termsFound(1, termEntry('や', 'や'));
      }
      if (text.startsWith('ほ')) {
        return termsFound(1, termEntry('帆', 'ほ', { originalText: 'ほ' }));
      }
      if (text.startsWith('ミナト')) {
        return termsFound(3, termEntry('ミナト', 'みなと'));
      }
      return null;
    },
  });

  const result = await requestYomitanScanTokens('やほっ ミナト', deps, quietLogger);

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

const RANKED_DICTIONARIES = ['JPDBv2㋕', 'Jiten', 'CC100'];

test('requestYomitanScanTokens extracts best frequency rank from selected termsFind entry', async () => {
  const deps = createBackendDeps({
    dictionaries: RANKED_DICTIONARIES,
    termsFind: (text) =>
      text.startsWith('潜み')
        ? termsFound(2, {
            ...termEntry('潜む', 'ひそむ', { originalText: '潜み' }),
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
              { headwordIndex: 0, dictionary: 'CC100', frequency: 118121, displayValue: null },
            ],
          })
        : null,
  });

  const result = await requestYomitanScanTokens('潜み', deps, quietLogger);

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

const LATER_EXACT_FREQUENCY_CASES = [
  {
    name: 'uses frequency from later exact-match entry when first exact entry has none',
    line: '者',
    later: {
      ...termEntry('者', 'もの'),
      frequencies: [
        { headwordIndex: 0, dictionary: 'JPDBv2㋕', frequency: 79601, displayValue: '475,79601句' },
        { headwordIndex: 0, dictionary: 'Jiten', frequency: 338, displayValue: '338' },
      ],
    },
    first: termEntry('者', 'もの'),
    expected: { surface: '者', reading: 'もの', headword: '者', headwordReading: 'もの' },
    expectedRank: 475,
  },
  {
    name: 'can use frequency from later exact secondary-match entry',
    line: '者',
    later: {
      ...termEntry('者', 'もの', { isPrimary: false }),
      frequencies: [
        { headwordIndex: 0, dictionary: 'JPDBv2㋕', frequency: 79601, displayValue: '475,79601句' },
      ],
    },
    first: termEntry('者', 'もの'),
    expected: { surface: '者', reading: 'もの', headword: '者', headwordReading: 'もの' },
    expectedRank: 475,
  },
  {
    name: 'uses exact frequency entry when selected reading differs',
    line: '第二走者',
    later: {
      ...termEntry('第二', '', { isPrimary: false }),
      frequencies: [
        {
          headwordIndex: 0,
          dictionary: 'JPDBv2㋕',
          frequency: 189513,
          displayValue: '1820,189513句',
        },
      ],
    },
    first: termEntry('第二', 'だいに'),
    expected: { surface: '第二', reading: 'だいに', headword: '第二', headwordReading: 'だいに' },
    expectedRank: 1820,
  },
];

for (const c of LATER_EXACT_FREQUENCY_CASES) {
  test(`requestYomitanScanTokens ${c.name}`, async () => {
    const surface = c.expected.surface;
    const deps = createBackendDeps({
      dictionaries: RANKED_DICTIONARIES,
      termsFind: (text) =>
        text.startsWith(surface)
          ? termsFound(surface.length, { ...c.first, frequencies: [] }, c.later)
          : null,
    });

    const result = await requestYomitanScanTokens(c.line, deps, quietLogger);

    assert.deepEqual(result?.[0], {
      ...c.expected,
      startPos: 0,
      endPos: surface.length,
      isNameMatch: false,
      frequencyRank: c.expectedRank,
    });
  });
}

test('requestYomitanScanTokens retries shorter windows when a greedy match has no exact-source headword', async () => {
  const deps = createBackendDeps({
    dictionaries: ['JMdict'],
    termsFind: (text) => {
      if (!text.startsWith('平')) {
        return null;
      }
      if (text.length >= 4) {
        // Simulates Yomitan normalization consuming punctuation/whitespace:
        // the greedy match spans 平 （平 but no headword source equals it.
        return termsFound(4, termEntry('平々', 'へいへい', { originalText: '平平' }));
      }
      return termsFound(1, termEntry('平', 'たいら'));
    },
  });

  const result = await requestYomitanScanTokens('平 （平）', deps, quietLogger);

  const token = {
    surface: '平',
    reading: 'たいら',
    headword: '平',
    headwordReading: 'たいら',
    isNameMatch: false,
    frequencyRank: undefined,
  };
  assert.deepEqual(result, [
    { ...token, startPos: 0, endPos: 1 },
    { ...token, startPos: 3, endPos: 4 },
  ]);
});

test('requestYomitanScanTokens emits complete readings for kanji-kana compounds', async () => {
  const deps = createBackendDeps({
    dictionaries: ['JPDBv2㋕'],
    termsFind: (text) =>
      text.startsWith('待ち合わせてる')
        ? termsFound(
            7,
            termEntry('待ち合わせる', 'まちあわせる', { originalText: '待ち合わせてる' }),
          )
        : null,
  });

  const result = await requestYomitanScanTokens('待ち合わせてる', deps, quietLogger);

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

function nameTokenSummary(result: Awaited<ReturnType<typeof requestYomitanScanTokens>>) {
  return result?.map(({ surface, headword, startPos, endPos, isNameMatch }) => ({
    surface,
    headword,
    startPos,
    endPos,
    isNameMatch,
  }));
}

test('requestYomitanScanTokens marks grouped entries when SubMiner dictionary alias only exists on definitions', async () => {
  const deps = createBackendDeps({
    termsFind: (text) =>
      text === 'カズマ'
        ? termsFound(3, {
            ...termEntry('カズマ', 'かずま'),
            dictionaryAlias: '',
            definitions: [
              { dictionary: 'JMdict', dictionaryAlias: 'JMdict' },
              { dictionary: CHARACTER_DICTIONARY, dictionaryAlias: CHARACTER_DICTIONARY },
            ],
          })
        : null,
  });

  const result = await requestYomitanScanTokens('カズマ', deps, quietLogger, {
    includeNameMatchMetadata: true,
  });

  assert.deepEqual(nameTokenSummary(result), [
    { surface: 'カズマ', headword: 'カズマ', startPos: 0, endPos: 3, isNameMatch: true },
  ]);
});

// A SubMiner character entry whose structured content carries `media` (an
// image path or a data attribute) naming the media it was generated for.
function characterEntryForMedia(term: string, reading: string, media: Record<string, unknown>) {
  return {
    ...termEntry(term, reading),
    definitions: [
      {
        dictionary: 'SubMiner Character Dictionary',
        dictionaryAlias: 'SubMiner Character Dictionary',
        entries: [{ type: 'structured-content', content: media }],
      },
    ],
  };
}

test('requestYomitanScanTokens ignores SubMiner character entries from other media', async () => {
  const deps = createBackendDeps({
    termsFind: (text) =>
      text === 'カズ'
        ? termsFound(
            2,
            characterEntryForMedia('カズ', 'かず', {
              tag: 'img',
              path: 'img/m115230-c9.png',
              alt: 'Kaz',
            }),
          )
        : null,
  });

  const result = await requestYomitanScanTokens('カズ', deps, quietLogger, {
    includeNameMatchMetadata: true,
    currentCharacterDictionaryMediaId: 21202,
  });

  // No dictionary-backed token survives (the only match belongs to another
  // media's character dictionary), so the line reports no tokenization.
  assert.equal(result, null);
});

test('requestYomitanScanTokens accepts SubMiner character entries with structured-content media data', async () => {
  const deps = createBackendDeps({
    termsFind: (text) =>
      text === 'アクア'
        ? termsFound(
            3,
            characterEntryForMedia('アクア', 'あくあ', {
              tag: 'div',
              data: { subminerMediaId: '21699' },
              content: [{ tag: 'img', path: 'img/m115230-c1.png', alt: 'アクア' }],
            }),
          )
        : null,
  });

  const result = await requestYomitanScanTokens('アクア', deps, quietLogger, {
    includeNameMatchMetadata: true,
    currentCharacterDictionaryMediaId: 21699,
  });

  assert.equal(result?.[0]?.surface, 'アクア');
  assert.equal(result?.[0]?.isNameMatch, true);
});

const NAME_DICTIONARIES = ['JMdict', CHARACTER_DICTIONARY];

function nameEntry(term: string, reading: string) {
  return termEntry(term, reading, { dictionary: CHARACTER_DICTIONARY });
}

function jmdictEntry(term: string, reading: string, originalText = term) {
  return termEntry(term, reading, { originalText, dictionary: 'JMdict' });
}

test('requestYomitanScanTokens greedily tokenizes character names before longer generic matches', async () => {
  const deps = createBackendDeps({
    dictionaries: NAME_DICTIONARIES,
    termsFind: (text) => {
      if (text.startsWith('美姫')) {
        return termsFound(2, nameEntry('美姫', 'みき'));
      }
      if (text.startsWith('とヨータ')) {
        // Greedy generic match: とヨー normalizes to とよう (渡洋). Without the
        // name pre-pass this consumes the ヨ of ヨータ.
        return termsFound(3, jmdictEntry('渡洋', 'とよう', 'とヨー'), jmdictEntry('と', 'と'));
      }
      if (text.startsWith('ヨータ')) {
        return termsFound(3, nameEntry('ヨータ', 'よーた'));
      }
      if (text === 'と') {
        return termsFound(1, jmdictEntry('と', 'と'));
      }
      return null;
    },
  });

  const result = await requestYomitanScanTokens('美姫とヨータ', deps, quietLogger, {
    includeNameMatchMetadata: true,
  });

  assert.deepEqual(nameTokenSummary(result), [
    { surface: '美姫', headword: '美姫', startPos: 0, endPos: 2, isNameMatch: true },
    { surface: 'と', headword: 'と', startPos: 2, endPos: 3, isNameMatch: false },
    { surface: 'ヨータ', headword: 'ヨータ', startPos: 3, endPos: 6, isNameMatch: true },
  ]);
});

test('requestYomitanScanTokens lets a longer generic word beat a shorter name at the same position', async () => {
  const deps = createBackendDeps({
    dictionaries: NAME_DICTIONARIES,
    termsFind: (text) => {
      if (text.startsWith('空気')) {
        // A character named 空 matches here, but the generic 空気 is longer and
        // must win the position.
        return termsFound(2, nameEntry('空', 'くう'), jmdictEntry('空気', 'くうき'));
      }
      if (text.startsWith('変わって')) {
        return termsFound(4, jmdictEntry('変わる', 'かわる', '変わって'));
      }
      return null;
    },
  });

  const result = await requestYomitanScanTokens('空気変わって', deps, quietLogger, {
    includeNameMatchMetadata: true,
  });

  assert.deepEqual(nameTokenSummary(result), [
    { surface: '空気', headword: '空気', startPos: 0, endPos: 2, isNameMatch: false },
    { surface: '変わって', headword: '変わる', startPos: 2, endPos: 6, isNameMatch: false },
  ]);
});

test('requestYomitanScanTokens lets a generic word beat a name it fully contains', async () => {
  const deps = createBackendDeps({
    dictionaries: NAME_DICTIONARIES,
    termsFind: (text) => {
      if (text.startsWith('写真')) {
        return termsFound(2, jmdictEntry('写真', 'しゃしん'));
      }
      if (text.startsWith('写')) {
        return termsFound(1, jmdictEntry('写', 'しゃ'));
      }
      if (text.startsWith('真')) {
        // The given name of 安田真 also matches the second half of 写真.
        return termsFound(1, nameEntry('真', 'しん'), jmdictEntry('真', 'しん'));
      }
      if (text.startsWith('は')) {
        return termsFound(1, jmdictEntry('は', 'は'));
      }
      return null;
    },
  });

  const result = await requestYomitanScanTokens('写真は', deps, quietLogger, {
    includeNameMatchMetadata: true,
  });

  assert.deepEqual(nameTokenSummary(result), [
    { surface: '写真', headword: '写真', startPos: 0, endPos: 2, isNameMatch: false },
    { surface: 'は', headword: 'は', startPos: 2, endPos: 3, isNameMatch: false },
  ]);
});

test('requestYomitanScanTokens still finds an emphatically elongated name a longer generic match would swallow', async () => {
  // Yomitan collapses emphatic sequences, so ミナァァト resolves to the ミナト
  // entry. The generic word とミナ starts earlier and would swallow the name
  // unless the pre-pass reserves it, so this only passes when the candidate
  // prefilter still treats the elongated spelling as a possible name start.
  const characterDictionary = 'SubMiner Character Dictionary (AniList 1)';
  const deps = createBackendDeps({
    dictionaries: ['JMdict', characterDictionary],
    termsFind: (text) => {
      if (text.startsWith('とミナ')) {
        return termsFound(3, jmdictEntry('トミナ', 'とみな', 'とミナ'));
      }
      if (text.startsWith('ミナァァト')) {
        return termsFound(
          5,
          termEntry('ミナト', 'みなと', {
            originalText: 'ミナァァト',
            dictionary: characterDictionary,
          }),
        );
      }
      if (text.startsWith('と')) {
        return termsFound(1, jmdictEntry('と', 'と'));
      }
      return null;
    },
  });

  const result = await requestYomitanScanTokens('とミナァァト', deps, quietLogger, {
    includeNameMatchMetadata: true,
    currentCharacterDictionaryMediaId: 1,
    nameCandidates: { key: 'media-1', forms: ['ミナト', 'みなと'] },
  });

  const nameToken = result?.find((token) => token.isNameMatch === true);
  assert.ok(nameToken, 'expected the elongated name to be reserved by the pre-pass');
  assert.equal(nameToken?.headword, 'ミナト');
  assert.equal(nameToken?.startPos, 1);
});

test('requestYomitanScanTokens preserves matched headword word classes', async () => {
  const deps = createBackendDeps({
    termsFind: (text) => {
      if (text !== 'は') {
        return null;
      }
      const entry = termEntry('は', 'は');
      return termsFound(1, {
        ...entry,
        headwords: [{ ...entry.headwords[0], wordClasses: ['prt'] }],
      });
    },
  });

  const result = await requestYomitanScanTokens('は', deps, quietLogger);

  assert.deepEqual(result?.[0]?.wordClasses, ['prt']);
});

test('requestYomitanScanTokens skips fallback fragments without exact primary source matches', async () => {
  const matches: Array<[string, ReturnType<typeof termsFound>]> = [
    ['だが ', termsFound(2, termEntry('だが', 'だが'))],
    ['それでも', termsFound(4, termEntry('それでも', 'それでも'))],
    ['届かぬ', termsFound(3, termEntry('届く', 'とどく', { originalText: '届かぬ' }))],
    ['高み', termsFound(2, termEntry('高み', 'たかみ'))],
    ['があった', termsFound(2, termEntry('があ', '', { originalText: 'が' }))],
    ['あった', termsFound(3, termEntry('ある', 'ある', { originalText: 'あった' }))],
  ];
  const deps = createBackendDeps({
    termsFind: (text) => matches.find(([prefix]) => text.startsWith(prefix))?.[1],
  });

  const result = await requestYomitanScanTokens(
    'だが それでも届かぬ高みがあった',
    deps,
    quietLogger,
  );

  assert.deepEqual(
    result?.map(({ surface, headword, startPos, endPos }) => ({
      surface,
      headword,
      startPos,
      endPos,
    })),
    [
      { surface: 'だが', headword: 'だが', startPos: 0, endPos: 2 },
      { surface: 'それでも', headword: 'それでも', startPos: 3, endPos: 7 },
      { surface: '届かぬ', headword: '届く', startPos: 7, endPos: 10 },
      { surface: '高み', headword: '高み', startPos: 10, endPos: 12 },
      // が has no exact primary source match, so it survives only as an
      // unparsed filler run (the parseText segmentation used to supply this).
      { surface: 'が', headword: 'が', startPos: 12, endPos: 13 },
      { surface: 'あった', headword: 'ある', startPos: 13, endPos: 16 },
    ],
  );
  assert.equal(result?.[4]?.isUnparsedRun, true);
});

function createCatScanDeps(lookups: string[]) {
  return createBackendDeps({
    lookups,
    termsFind: (text) => (text.startsWith('猫') ? termsFound(1, termEntry('猫', 'ねこ')) : null),
  });
}

test('requestYomitanScanTokens reuses the cross-line termsFind cache for repeated lookups', async () => {
  const lookups: string[] = [];
  const deps = createCatScanDeps(lookups);

  const first = await requestYomitanScanTokens('猫', deps, quietLogger);
  const second = await requestYomitanScanTokens('猫', deps, quietLogger);

  assert.equal(first?.length, 1);
  assert.equal(second?.length, 1);
  // The second line hits the window-persistent cache: no new backend lookup.
  assert.equal(countTermsFindLookups(lookups, '猫'), 1);
});

test('clearYomitanParserCachesForWindow invalidates the cross-line termsFind cache', async () => {
  const lookups: string[] = [];
  const deps = createCatScanDeps(lookups);

  await requestYomitanScanTokens('猫', deps, quietLogger);
  clearYomitanParserCachesForWindow(deps.getYomitanParserWindow() as never);
  await requestYomitanScanTokens('猫', deps, quietLogger);

  assert.equal(countTermsFindLookups(lookups, '猫'), 2);
});

test('an oversized termsFind result is dropped from the cache instead of being reused', async () => {
  const lookups: string[] = [];
  // One entry over the runtime's 20,000 retained-entry budget: the weight is
  // only known once the lookup resolves, so the cache has to re-check then.
  const oversizedEntries = Array.from({ length: 20_001 }, () => termEntry('猫', 'ねこ'));
  const deps = createBackendDeps({
    lookups,
    termsFind: (text) => (text.startsWith('猫') ? termsFound(1, ...oversizedEntries) : null),
  });

  await requestYomitanScanTokens('猫', deps, quietLogger);
  await requestYomitanScanTokens('猫', deps, quietLogger);

  assert.equal(countTermsFindLookups(lookups, '猫'), 2);
});

// A window that "matches" its whole length but never yields an exact-source
// headword: the worst case for the retry ladder.
function mismatchAcross(length: number) {
  return termsFound(length, termEntry('ミスマッチ', 'みすまっち', { originalText: 'ZZZ' }));
}

test('scanner tokens survive a retry-budget escalation whose parseText finds nothing', async () => {
  const parsedTexts: string[] = [];
  const deps = createBackendDeps({
    // The rest of the line burns the blind-retry budget at every position.
    termsFind: (text) =>
      text.startsWith('猫') ? termsFound(1, termEntry('猫', 'ねこ')) : mismatchAcross(text.length),
    actions: {
      parseText: (params) => {
        parsedTexts.push((params as { text: string }).text);
        return [];
      },
    },
  });

  const result = await requestYomitanScanTokens('猫あいうえおかきくけこ', deps, quietLogger);

  // The escalation ran exactly once and found nothing, so the tokens the
  // scanner did resolve are kept instead of dropping the line to raw text.
  assert.deepEqual(parsedTexts, ['猫あいうえおかきくけこ']);
  assert.equal(result?.[0]?.surface, '猫');
});

test('requestYomitanScanTokens skips termsFind lookups at punctuation and whitespace positions', async () => {
  const lookups: string[] = [];
  const deps = createCatScanDeps(lookups);

  const result = await requestYomitanScanTokens('「猫」…♪', deps, quietLogger);

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
  const deps = createBackendDeps({
    lookups,
    // Each step down the ladder is a blind guess with nothing shorter
    // reported to aim at.
    termsFind: (text) => mismatchAcross(text.length),
    actions: {
      parseText: (params) => {
        parsedTexts.push((params as { text: string }).text);
        return [
          {
            source: 'scanning-parser',
            index: 0,
            content: [
              [{ text: 'あいうえお', reading: 'あいうえお', headwords: [[{ term: 'あい' }]] }],
            ],
          },
        ];
      },
    },
  });

  const result = await requestYomitanScanTokens('あいうえおかきくけこ', deps, quietLogger);

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
  const deps = createBackendDeps({
    lookups,
    // Normalization keeps eating one character past the term, so every window
    // reports a shorter consumed length: informative steps that must not be
    // spent from the blind-retry budget. The term only surfaces at length 2,
    // six lookups down the ladder.
    termsFind: (text) =>
      text.length === 2
        ? termsFound(2, termEntry('あい', 'あい'))
        : mismatchAcross(Math.max(text.length - 1, 0)),
  });

  const result = await requestYomitanScanTokens('あいうえおかきくけこさしすせ', deps, quietLogger);

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

test('getYomitanDictionaryInfo normalizes the backend dictionary list', async () => {
  const deps = createBackendDeps({
    dictionaryInfo: [
      { title: ` ${CHARACTER_DICTIONARY} `, revision: '1', frequencyMode: 'rank-based' },
      { title: 'JPDB', revision: 3, frequencyMode: 'unknown' },
      { title: '   ' },
      'not-an-entry',
    ],
  });

  const dictionaries = await getYomitanDictionaryInfo(deps, quietLogger);

  assert.deepEqual(dictionaries, [
    { title: CHARACTER_DICTIONARY, revision: '1', frequencyMode: 'rank-based' },
    { title: 'JPDB', revision: 3, frequencyMode: undefined },
  ]);
});

test('dictionary settings helpers upsert and remove dictionary entries without reordering', async () => {
  const title = 'SubMiner Character Dictionary (AniList 1)';
  const optionsFull = {
    profileCurrent: 0,
    profiles: [
      {
        options: {
          dictionaries: [
            { name: 'Jitendex', alias: 'Jitendex', enabled: true },
            { name: title, alias: title, enabled: false },
          ],
        },
      },
    ],
  };
  const saved: Array<typeof optionsFull> = [];
  const deps = createBackendDeps({
    actions: {
      optionsGetFull: () => structuredClone(optionsFull),
      setAllSettings: (params) => {
        saved.push((params as { value: typeof optionsFull }).value);
        return true;
      },
    },
  });

  const upserted = await upsertYomitanDictionarySettings(title, 'all', deps, quietLogger);
  const removed = await removeYomitanDictionarySettings(title, 'all', 'delete', deps, quietLogger);

  assert.equal(upserted, true);
  assert.equal(removed, true);
  assert.deepEqual(
    saved.map((value) =>
      value.profiles[0]?.options.dictionaries.map(({ name, enabled }) => ({ name, enabled })),
    ),
    [
      [
        { name: 'Jitendex', enabled: true },
        { name: title, enabled: true },
      ],
      [{ name: 'Jitendex', enabled: true }],
    ],
  );
});

function writeTempDictionaryZip(): string {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-yomitan-import-'));
  const zipPath = path.join(tempDir, 'dict.zip');
  fs.writeFileSync(zipPath, Buffer.from('zip-bytes'));
  return zipPath;
}

test('importYomitanDictionaryFromZip imports via localhost URL instead of embedding archive bytes in script', async () => {
  const zipPath = writeTempDictionaryZip();
  const servedArchives: string[] = [];
  const base64Imports: string[] = [];
  const deps = createSettingsAutomationDeps({
    importDictionaryArchiveUrl: async (url: string) => {
      servedArchives.push(await (await fetch(url)).text());
    },
    importDictionaryArchiveBase64: async (archive: string) => {
      base64Imports.push(archive);
    },
  });

  const imported = await importYomitanDictionaryFromZip(zipPath, deps, quietLogger);

  assert.equal(imported, true);
  assert.deepEqual(servedArchives, ['zip-bytes']);
  assert.deepEqual(base64Imports, []);
});

test('importYomitanDictionaryFromZip falls back to base64 import for older Yomitan bridge', async () => {
  const zipPath = writeTempDictionaryZip();
  const base64Imports: Array<{ archive: string; fileName: string }> = [];
  const deps = createSettingsAutomationDeps({
    importDictionaryArchiveBase64: async (archive: string, fileName: string) => {
      base64Imports.push({ archive: Buffer.from(archive, 'base64').toString(), fileName });
    },
  });

  const imported = await importYomitanDictionaryFromZip(zipPath, deps, quietLogger);

  assert.equal(imported, true);
  assert.deepEqual(base64Imports, [{ archive: 'zip-bytes', fileName: 'dict.zip' }]);
});

test('importYomitanDictionaryFromZip returns false when served archive cannot be read', async () => {
  const zipPath = writeTempDictionaryZip();
  const deps = createSettingsAutomationDeps({
    importDictionaryArchiveUrl: async (url: string) => {
      fs.unlinkSync(zipPath);
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`archive fetch failed: ${response.status}`);
      }
    },
  });

  const imported = await importYomitanDictionaryFromZip(zipPath, deps, quietLogger);

  assert.equal(imported, false);
});

test('deleteYomitanDictionaryByTitle deletes through the settings automation bridge', async () => {
  const deletedTitles: string[] = [];
  const deps = createSettingsAutomationDeps({
    deleteDictionary: async (title: string) => {
      deletedTitles.push(title);
    },
  });

  const deleted = await deleteYomitanDictionaryByTitle(
    ` ${CHARACTER_DICTIONARY} `,
    deps,
    quietLogger,
  );

  assert.equal(deleted, true);
  assert.deepEqual(deletedTitles, [CHARACTER_DICTIONARY]);
});

test('addYomitanNoteViaSearch returns note and duplicate ids from the bridge payload', async () => {
  const deps = createDeps(async (_script) => ({
    noteId: 42,
    duplicateNoteIds: [18, 7, 18],
  }));

  const result = await addYomitanNoteViaSearch('食べる', deps, quietLogger);

  assert.deepEqual(result, {
    noteId: 42,
    duplicateNoteIds: [18, 7, 18],
  });
});

test('addYomitanNoteViaSearch rejects invalid numeric note ids from the bridge shortcut', async () => {
  const deps = createDeps(async () => NaN);

  const result = await addYomitanNoteViaSearch('食べる', deps, quietLogger);

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

  const result = await addYomitanNoteViaSearch('食べる', deps, quietLogger);

  assert.deepEqual(result, {
    noteId: null,
    duplicateNoteIds: [18, 7],
  });
});
