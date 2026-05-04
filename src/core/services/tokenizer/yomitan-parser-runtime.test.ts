import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import * as vm from 'node:vm';
import {
  addYomitanNoteViaSearch,
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

async function runInjectedYomitanScript(
  script: string,
  handler: (action: string, params: unknown) => unknown,
): Promise<unknown> {
  return await vm.runInNewContext(script, {
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
  });
}

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

test('requestYomitanScanTokens uses left-to-right termsFind scanning instead of parseText', async () => {
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
            },
          },
        ],
      };
    }
    return [
      {
        surface: 'カズマ',
        reading: 'かずま',
        headword: 'カズマ',
        startPos: 0,
        endPos: 3,
      },
    ];
  });

  const result = await requestYomitanScanTokens('カズマ', deps, {
    error: () => undefined,
  });

  assert.deepEqual(result, [
    {
      surface: 'カズマ',
      reading: 'かずま',
      headword: 'カズマ',
      startPos: 0,
      endPos: 3,
    },
  ]);
  const scannerScript = scripts.find((script) => script.includes('termsFind'));
  assert.ok(scannerScript, 'expected termsFind scanning request script');
  assert.doesNotMatch(scannerScript ?? '', /parseText/);
  assert.match(scannerScript ?? '', /matchType:\s*"exact"/);
  assert.match(scannerScript ?? '', /deinflect:\s*true/);
});

test('requestYomitanScanTokens extracts best frequency rank from selected termsFind entry', async () => {
  let scannerScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('termsFind')) {
      scannerScript = script;
      return [];
    }
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profileIndex: 0,
        scanLength: 40,
        dictionaries: ['JPDBv2㋕', 'Jiten', 'CC100'],
        dictionaryPriorityByName: {
          'JPDBv2㋕': 0,
          Jiten: 1,
          CC100: 2,
        },
        dictionaryFrequencyModeByName: {
          'JPDBv2㋕': 'rank-based',
          Jiten: 'rank-based',
          CC100: 'rank-based',
        },
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
    return null;
  });

  await requestYomitanScanTokens('潜み', deps, {
    error: () => undefined,
  });

  const result = await runInjectedYomitanScript(scannerScript, (action, params) => {
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

  assert.deepEqual(result, [
    {
      surface: '潜み',
      reading: 'ひそ',
      headword: '潜む',
      startPos: 0,
      endPos: 2,
      isNameMatch: false,
      frequencyRank: 4073,
    },
  ]);
});

test('requestYomitanScanTokens uses frequency from later exact-match entry when first exact entry has none', async () => {
  let scannerScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('termsFind')) {
      scannerScript = script;
      return [];
    }
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profileIndex: 0,
        scanLength: 40,
        dictionaries: ['JPDBv2㋕', 'Jiten', 'CC100'],
        dictionaryPriorityByName: {
          'JPDBv2㋕': 0,
          Jiten: 1,
          CC100: 2,
        },
        dictionaryFrequencyModeByName: {
          'JPDBv2㋕': 'rank-based',
          Jiten: 'rank-based',
          CC100: 'rank-based',
        },
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
    return null;
  });

  await requestYomitanScanTokens('者', deps, {
    error: () => undefined,
  });

  const result = await runInjectedYomitanScript(scannerScript, (action, params) => {
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

  assert.deepEqual(result, [
    {
      surface: '者',
      reading: 'もの',
      headword: '者',
      startPos: 0,
      endPos: 1,
      isNameMatch: false,
      frequencyRank: 475,
    },
  ]);
});

test('requestYomitanScanTokens can use frequency from later exact secondary-match entry', async () => {
  let scannerScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('termsFind')) {
      scannerScript = script;
      return [];
    }
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profileIndex: 0,
        scanLength: 40,
        dictionaries: ['JPDBv2㋕', 'Jiten', 'CC100'],
        dictionaryPriorityByName: {
          'JPDBv2㋕': 0,
          Jiten: 1,
          CC100: 2,
        },
        dictionaryFrequencyModeByName: {
          'JPDBv2㋕': 'rank-based',
          Jiten: 'rank-based',
          CC100: 'rank-based',
        },
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
    return null;
  });

  await requestYomitanScanTokens('者', deps, {
    error: () => undefined,
  });

  const result = await runInjectedYomitanScript(scannerScript, (action, params) => {
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

  assert.deepEqual(result, [
    {
      surface: '者',
      reading: 'もの',
      headword: '者',
      startPos: 0,
      endPos: 1,
      isNameMatch: false,
      frequencyRank: 475,
    },
  ]);
});

test('requestYomitanScanTokens uses exact frequency entry when selected reading differs', async () => {
  let scannerScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('termsFind')) {
      scannerScript = script;
      return [];
    }
    if (script.includes('optionsGetFull')) {
      return {
        profileCurrent: 0,
        profileIndex: 0,
        scanLength: 40,
        dictionaries: ['JPDBv2㋕', 'Jiten', 'CC100'],
        dictionaryPriorityByName: {
          'JPDBv2㋕': 0,
          Jiten: 1,
          CC100: 2,
        },
        dictionaryFrequencyModeByName: {
          'JPDBv2㋕': 'rank-based',
          Jiten: 'rank-based',
          CC100: 'rank-based',
        },
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
    return null;
  });

  await requestYomitanScanTokens('第二走者', deps, {
    error: () => undefined,
  });

  const result = (await runInjectedYomitanScript(scannerScript, (action, params) => {
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
  })) as Array<Record<string, unknown>>;

  assert.deepEqual(result?.[0], {
    surface: '第二',
    reading: 'だいに',
    headword: '第二',
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
  let scannerScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('termsFind')) {
      scannerScript = script;
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
  assert.match(scannerScript, /const includeNameMatchMetadata = false;/);
});

test('requestYomitanScanTokens marks grouped entries when SubMiner dictionary alias only exists on definitions', async () => {
  let scannerScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('termsFind')) {
      scannerScript = script;
      return [];
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
    return null;
  });

  await requestYomitanScanTokens(
    'カズマ',
    deps,
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  assert.match(scannerScript, /getPreferredHeadword/);

  const result = await runInjectedYomitanScript(scannerScript, (action, params) => {
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
  });

  assert.equal(Array.isArray(result), true);
  assert.equal((result as { length?: number } | null)?.length, 1);
  assert.equal((result as Array<{ surface?: string }>)[0]?.surface, 'カズマ');
  assert.equal((result as Array<{ headword?: string }>)[0]?.headword, 'カズマ');
  assert.equal((result as Array<{ startPos?: number }>)[0]?.startPos, 0);
  assert.equal((result as Array<{ endPos?: number }>)[0]?.endPos, 3);
  assert.equal((result as Array<{ isNameMatch?: boolean }>)[0]?.isNameMatch, true);
});

test('requestYomitanScanTokens preserves matched headword word classes', async () => {
  let scannerScript = '';
  const deps = createDeps(async (script) => {
    if (script.includes('termsFind')) {
      scannerScript = script;
      return [];
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
    return null;
  });

  await requestYomitanScanTokens('は', deps, { error: () => undefined });

  const result = await runInjectedYomitanScript(scannerScript, (action, params) => {
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

  assert.deepEqual((result as Array<{ wordClasses?: string[] }>)[0]?.wordClasses, ['prt']);
});

test('requestYomitanScanTokens skips fallback fragments without exact primary source matches', async () => {
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

    return await runInjectedYomitanScript(script, (action, params) => {
      if (action !== 'termsFind') {
        throw new Error(`unexpected action: ${action}`);
      }

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
    });
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
      {
        surface: 'あった',
        headword: 'ある',
        startPos: 13,
        endPos: 16,
      },
    ],
  );
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

test('importYomitanDictionaryFromZip uses settings automation bridge instead of custom backend action', async () => {
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
    scripts.some((script) => script.includes('importDictionaryArchiveBase64')),
    true,
  );
  assert.equal(
    scripts.some((script) => script.includes('subminerImportDictionary')),
    false,
  );
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
