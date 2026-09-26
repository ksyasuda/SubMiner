import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { pathToFileURL } from 'node:url';
import * as vm from 'node:vm';
import { HACHIDORI_PARSER_BRIDGE_SCRIPT } from './hachidori-parser-bridge';
import {
  requestYomitanScanTokens,
  requestYomitanParseResults,
  requestYomitanTermFrequencies,
  syncYomitanDefaultAnkiServer,
  addYomitanNoteViaSearch,
} from './yomitan-parser-runtime';
import { createDeps } from './yomitan-scan-test-harness';
import { selectYomitanParseTokens } from './parser-selection-stage';

const extensionPath = path.resolve(__dirname, '../../../../vendor/hachidori/extension');
const characterDictionary = 'SubMiner Character Dictionary (AniList 1)';

function lookupResult(matched: string, expression: string, reading: string, dictionary = 'JMdict') {
  return {
    matched,
    deinflected: expression,
    trace: [],
    term: {
      expression,
      reading,
      score: 0,
      rules: 'v1',
      glossaries: [{ dictionary, glossary: '["definition"]', termTags: '', definitionTags: '' }],
      pitches: [],
      frequencies: [{ dictionary: 'Frequency', frequencies: [{ value: 42, displayValue: '' }] }],
    },
  };
}

async function createHarness(emptyLibrary = false) {
  const messages: Array<Record<string, unknown>> = [];
  let dictionaryRevision = 2;
  let optionRevision = 3;
  let anki = {
    url: 'http://127.0.0.1:8765',
    templates: [
      { id: 'default', name: 'Default', deck: 'Old', model: 'Japanese', fields: {} },
      { id: 'second', name: 'Second', deck: 'Other', model: 'Japanese', fields: {} },
    ],
  };
  let dictionaries = [
    { id: 'terms', title: 'JMdict', displayName: 'Main dictionary', enabled: true, revision: '1' },
    { id: 'names', title: characterDictionary, enabled: true, revision: '1' },
    {
      id: 'frequency',
      title: 'Frequency',
      enabled: true,
      revision: '1',
      frequencyMode: 'rank-based',
      frequencyCount: 1,
    },
  ];
  let duplicate = false;
  let proxyUrl: unknown;
  let loadingStatusReplies = 0;
  let busyEngineReplies = 0;
  const apiModule: unknown = await import(
    pathToFileURL(path.join(extensionPath, 'api-host.js')).href
  );
  const context = vm.createContext({
    __testApiModule: apiModule,
    setTimeout,
    crypto,
    URL,
    Blob,
    Uint8Array,
    atob,
    chrome: {
      storage: {
        local: {
          set: async (value: Record<string, unknown>) => {
            proxyUrl = value.subminerAnkiProxyUrl;
          },
          get: async () => ({
            options: { anki, scanLength: 40, revision: optionRevision },
            subminerAnkiProxyUrl: proxyUrl,
          }),
        },
      },
      runtime: {
        getManifest: () => ({ version: 'test' }),
        sendMessage: async (message: Record<string, unknown>) => {
          messages.push(structuredClone(message));
          if (message.type === 'hd_status') {
            const loading = loadingStatusReplies > 0;
            if (loading) loadingStatusReplies -= 1;
            return { ok: true, ready: !loading, loading };
          }
          if (busyEngineReplies > 0 && message.target !== 'hachidori-anki') {
            busyEngineReplies -= 1;
            return { ok: false, error: 'the dictionary engine is busy mutating' };
          }
          switch (message.type) {
            case 'hd_state_read':
              return {
                ok: true,
                state: emptyLibrary ? null : { revision: dictionaryRevision, dictionaries },
              };
            case 'hd_lookup': {
              const text = String(message.text);
              const candidates = [
                lookupResult('食べた', '食べる', 'たべる'),
                lookupResult('食べる', '食べる', 'たべる'),
                lookupResult('ミナト', 'ミナト', 'みなと', characterDictionary),
              ];
              return {
                ok: true,
                generation: 1,
                results: candidates.filter((result) => text.startsWith(result.matched)),
              };
            }
            case 'hd_options_write': {
              if (message.baseRevision !== optionRevision) return { ok: false, error: 'conflict' };
              const update = message.options;
              assert.ok(update && typeof update === 'object' && 'anki' in update);
              const value = update.anki;
              assert.ok(
                value &&
                  typeof value === 'object' &&
                  'url' in value &&
                  typeof value.url === 'string',
              );
              assert.ok('templates' in value && Array.isArray(value.templates));
              anki = { url: value.url, templates: value.templates };
              optionRevision += 1;
              return { ok: true };
            }
            case 'hd_apply_state': {
              assert.equal(message.baseRevision, dictionaryRevision);
              assert.ok(Array.isArray(message.dictionaries));
              dictionaries = message.dictionaries;
              dictionaryRevision += 1;
              return { ok: true };
            }
            case 'hd_anki_status':
              return { ok: true, configKey: 'configuration' };
            case 'hd_anki_preflight':
              return {
                ok: true,
                canAdd: !duplicate,
                state: duplicate ? 'duplicate' : 'addable',
                noteIds: duplicate ? [15] : [],
              };
            case 'hd_anki_submit':
              return { ok: true, state: 'added', noteId: 19 };
            case 'hd_import':
              return { ok: true, report: { success: true } };
            case 'hd_remove':
              return { ok: true };
            default:
              throw new Error('Unexpected native request: ' + String(message.type));
          }
        },
      },
    },
  });
  vm.runInContext('globalThis.window = globalThis', context);
  vm.runInContext(await readFile(path.join(extensionPath, 'reader-options.js'), 'utf8'), context);
  const script = HACHIDORI_PARSER_BRIDGE_SCRIPT.replace(
    "await import('./api-host.js')",
    '__testApiModule',
  ).replace("await import('./reader-options.js')", 'Promise.resolve()');
  await vm.runInContext(script, context);
  const run = async (code: string): Promise<unknown> =>
    structuredClone(await vm.runInContext(code, context));
  const invoke = (action: string, params?: unknown) =>
    run(`new Promise((resolve, reject) => {
    __subminerDictionarySendMessage(${JSON.stringify({ action, params })}, response => {
      if (response.error) reject(new Error(response.error.message)); else resolve(response.result);
    });
  })`);
  return {
    deps: createDeps(run),
    run,
    invoke,
    messages,
    anki: () => anki,
    proxyUrl: () => proxyUrl,
    setAnkiServer: (url: string) => {
      anki = { ...anki, url };
    },
    setAnkiTemplates: (templates: typeof anki.templates) => {
      anki = { ...anki, templates };
    },
    dictionaries: () => dictionaries,
    disableDictionary: (id: string) => {
      dictionaries = dictionaries.map((entry) =>
        entry.id === id ? { ...entry, enabled: false } : entry,
      );
      dictionaryRevision += 1;
    },
    setDuplicate: () => {
      duplicate = true;
    },
    changeRevision: () => {
      optionRevision += 1;
    },
    setLoadingStatusReplies: (count: number) => {
      loadingStatusReplies = count;
    },
    setBusyEngineReplies: (count: number) => {
      busyEngineReplies = count;
    },
  };
}

test('Hachidori runs the shared scanner with inflected offsets, headwords, names and frequencies', async () => {
  const harness = await createHarness();
  const tokens = await requestYomitanScanTokens(
    'ミナト 食べた',
    harness.deps,
    { error: assert.fail },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
    },
  );
  assert.ok(tokens);
  assert.equal(tokens[0]?.surface, 'ミナト');
  assert.equal(tokens[0]?.isNameMatch, true);
  assert.equal(tokens[1]?.surface, '食べた');
  assert.equal(tokens[1]?.headword, '食べる');
  assert.equal(tokens[1]?.startPos, 4);
  assert.equal(tokens[1]?.endPos, 7);
  assert.equal(tokens[1]?.frequencyRank, 42);
  assert.deepEqual(tokens[1]?.wordClasses, ['v1']);
  const frequencies = await requestYomitanTermFrequencies(
    [{ term: '食べる', reading: 'たべる' }],
    harness.deps,
    { error: assert.fail },
  );
  assert.equal(frequencies[0]?.frequency, 42);
  assert.equal(frequencies[0]?.dictionary, 'Frequency');
});

test('Hachidori frequency lookups match API headwords, readings and requested dictionaries', async () => {
  const harness = await createHarness();
  const query = (term: string, reading: string | null, dictionaries = ['Frequency']) =>
    harness.invoke('getTermFrequencies', { termReadingList: [{ term, reading }], dictionaries });
  assert.deepEqual(await query('食べる', 'たべる'), [
    {
      term: '食べる',
      reading: 'たべる',
      hasReading: false,
      dictionary: 'Frequency',
      frequency: 42,
      displayValue: null,
      displayValueParsed: false,
    },
  ]);
  assert.deepEqual(await query('食べる', 'べつのよみ'), []);
  assert.deepEqual(await query('食べる', null, ['Other frequency']), []);
  assert.deepEqual(await query('食べるだけ', null), []);
  assert.deepEqual(await query('頻度だけ', null), []);
  assert.deepEqual(await query('食べる', null), await query('食べる', 'たべる'));
});

test('Hachidori syncs the Anki endpoint and every term template through revisioned writes', async () => {
  const harness = await createHarness();
  const synced = await syncYomitanDefaultAnkiServer(
    'http://127.0.0.1:8766',
    harness.deps,
    { error: assert.fail },
    { deck: 'Mining' },
  );
  assert.equal(synced, true);
  assert.equal(harness.anki().url, 'http://127.0.0.1:8766');
  assert.deepEqual(
    harness.anki().templates.map((template) => template.deck),
    ['Mining', 'Mining'],
  );
  assert.equal(harness.proxyUrl(), null);
  assert.equal(
    await syncYomitanDefaultAnkiServer(
      'http://127.0.0.1:8766',
      harness.deps,
      { error: assert.fail },
      { forceOverride: true },
    ),
    true,
  );
  assert.equal(harness.proxyUrl(), 'http://127.0.0.1:8766');
  await assert.rejects(
    harness.invoke('setAllSettings', {
      value: {
        hachidoriRevisions: { dictionaries: 99, options: 99 },
        profiles: [{ options: { dictionaries: [], anki: { server: '', cardFormats: [] } } }],
      },
    }),
    /settings changed/,
  );
});

test('Hachidori applies only SubMiner changes when its settings moved on meanwhile', async () => {
  const harness = await createHarness();
  const projected = (await harness.invoke('optionsGetFull')) as {
    profiles: Array<{ options: { dictionaries: Array<{ name: string; enabled: boolean }> } }>;
  };
  const jmdict = projected.profiles[0]!.options.dictionaries.find(
    (entry) => entry.name === 'JMdict',
  );
  jmdict!.enabled = false;
  // Hachidori changed both revisions after SubMiner read them.
  harness.disableDictionary('names');
  harness.setAnkiServer('http://192.168.1.10:8765');
  harness.changeRevision();
  assert.equal(await harness.invoke('setAllSettings', { value: projected }), true);
  assert.deepEqual(
    harness.dictionaries().map((entry) => [entry.id, entry.enabled]),
    [
      ['terms', false],
      ['names', false],
      ['frequency', true],
    ],
  );
  assert.equal(harness.anki().url, 'http://192.168.1.10:8765');
  assert.equal(harness.messages.filter((message) => message.type === 'hd_options_write').length, 0);
});

test('Hachidori settings writes survive an empty template list', async () => {
  const harness = await createHarness();
  harness.setAnkiTemplates([]);
  assert.equal(
    await syncYomitanDefaultAnkiServer(
      'http://127.0.0.1:8766',
      harness.deps,
      { error: assert.fail },
      {
        forceOverride: true,
      },
    ),
    true,
  );
  assert.equal(harness.anki().url, 'http://127.0.0.1:8766');
});

test('Hachidori initializes an empty library before the first dictionary import', async () => {
  const harness = await createHarness(true);
  assert.deepEqual(await harness.invoke('getDictionaryInfo'), []);
  assert.equal(
    await syncYomitanDefaultAnkiServer(
      'http://127.0.0.1:8766',
      harness.deps,
      { error: assert.fail },
      { forceOverride: true },
    ),
    true,
  );
});

test('Hachidori restores direct AnkiConnect when disabling its managed proxy', async () => {
  const harness = await createHarness();
  const logger = { error: assert.fail };
  assert.equal(
    await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', harness.deps, logger, {
      forceOverride: true,
    }),
    true,
  );
  assert.equal(harness.anki().url, 'http://127.0.0.1:8766');
  assert.equal(
    await syncYomitanDefaultAnkiServer('http://127.0.0.1:8765', harness.deps, logger),
    true,
  );
  assert.equal(harness.anki().url, 'http://127.0.0.1:8765');
  assert.equal(harness.proxyUrl(), null);
});

test('Hachidori preserves a custom Anki endpoint when disabling its managed proxy', async () => {
  const harness = await createHarness();
  const logger = { error: assert.fail };
  assert.equal(
    await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', harness.deps, logger, {
      forceOverride: true,
    }),
    true,
  );
  harness.setAnkiServer('http://192.168.1.10:8765');
  assert.equal(
    await syncYomitanDefaultAnkiServer('http://127.0.0.1:8765', harness.deps, logger),
    false,
  );
  assert.equal(harness.anki().url, 'http://192.168.1.10:8765');
  assert.equal(harness.proxyUrl(), null);
});

test('Hachidori fallback parsing retains token boundaries, inflected readings and headwords', async () => {
  const harness = await createHarness();
  const parsed = await requestYomitanParseResults('ミナト 食べた', harness.deps, {
    error: assert.fail,
  });
  const tokens = selectYomitanParseTokens(parsed, () => false, 'headword');
  assert.deepEqual(
    tokens?.map((token) => ({
      surface: token.surface,
      headword: token.headword,
      reading: token.reading,
      start: token.startPos,
    })),
    [
      { surface: 'ミナト', headword: 'ミナト', reading: 'みなと', start: 0 },
      { surface: '食べた', headword: '食べる', reading: 'たべた', start: 4 },
    ],
  );
});

test('Hachidori stats mining returns note IDs and prevents duplicate submissions', async () => {
  const harness = await createHarness();
  assert.deepEqual(await addYomitanNoteViaSearch('食べる', harness.deps, { error: assert.fail }), {
    noteId: 19,
    duplicateNoteIds: [],
  });
  harness.setDuplicate();
  assert.deepEqual(await addYomitanNoteViaSearch('食べる', harness.deps, { error: assert.fail }), {
    noteId: null,
    duplicateNoteIds: [15],
  });
  assert.equal(harness.messages.filter((message) => message.type === 'hd_anki_submit').length, 1);
});

test('Hachidori mining requests render native dictionary aliases and frequency markers', async () => {
  const harness = await createHarness();
  await addYomitanNoteViaSearch('食べる', harness.deps, { error: assert.fail });
  const request = harness.messages.find((message) => message.type === 'hd_anki_preflight')?.request;
  assert.ok(request && typeof request === 'object');
  assert.ok('subminerEnrich' in request && request.subminerEnrich === false);
  const native: unknown = await import(
    pathToFileURL(path.join(extensionPath, 'anki-values.js')).href
  );
  assert.ok(native && typeof native === 'object' && 'buildAnkiFields' in native);
  assert.equal(typeof native.buildAnkiFields, 'function');
  if (typeof native.buildAnkiFields !== 'function') assert.fail('Native renderer is unavailable');
  const fields: unknown = await native.buildAnkiFields(
    request,
    {
      Dictionary: { value: '{dictionary-alias}' },
      Frequency: { value: '{single-frequency-frequency}' },
      Rank: { value: '{frequency-harmonic-rank}' },
    },
    {},
  );
  assert.deepEqual(fields, {
    Dictionary: 'Main dictionary',
    Frequency: '<ul style="text-align: left;"><li>Frequency: </li></ul>',
    Rank: '42',
  });
  assert.ok('dictionaryIds' in request);
  assert.deepEqual(request.dictionaryIds, {
    JMdict: 'terms',
    [characterDictionary]: 'names',
    Frequency: 'frequency',
  });
  assert.deepEqual(
    harness.messages.find((message) => message.type === 'hd_anki_submit')?.request,
    request,
  );
});

test('Hachidori settings automation imports ZIP bytes and removes the matching dictionary ID', async () => {
  const harness = await createHarness();
  await harness.run(
    "__subminerYomitanSettingsAutomation.importDictionaryArchiveBase64('UEs=', 'characters.zip')",
  );
  await harness.run(
    `__subminerYomitanSettingsAutomation.deleteDictionary(${JSON.stringify(characterDictionary)})`,
  );
  const imported = harness.messages.find((message) => message.type === 'hd_import');
  assert.equal(imported?.fileName, 'characters.zip');
  assert.match(String(imported?.blobUrl), /^blob:/);
  const removed = harness.messages.find((message) => message.type === 'hd_remove');
  assert.equal(removed?.id, 'names');
});

test('Hachidori waits for a loading engine and retries busy requests before answering', async () => {
  const harness = await createHarness();
  harness.setLoadingStatusReplies(2);
  harness.setBusyEngineReplies(1);
  const dictionaries = (await harness.invoke('getDictionaryInfo')) as Array<{ title: string }>;
  assert.equal(dictionaries.length, 3);
  const statusCalls = harness.messages.filter((message) => message.type === 'hd_status').length;
  assert.ok(statusCalls >= 3, `expected repeated status polls, saw ${statusCalls}`);
  const tokens = await requestYomitanScanTokens('食べた', harness.deps, { error: assert.fail });
  assert.equal(tokens?.[0]?.headword, '食べる');
  assert.equal(tokens?.[0]?.frequencyRank, 42);
});
