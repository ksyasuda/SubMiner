import assert from 'node:assert/strict';
import test from 'node:test';
import {
  requestYomitanTermFrequencies,
  syncYomitanDefaultAnkiServer,
} from './yomitan-parser-runtime';

function createDeps(executeJavaScript: (script: string) => Promise<unknown>) {
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
  };
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
  assert.equal(infoLogs.length, 1);
});

test('syncYomitanDefaultAnkiServer returns false when script reports no change', async () => {
  const deps = createDeps(async () => ({ updated: false }));

  const updated = await syncYomitanDefaultAnkiServer('http://127.0.0.1:8766', deps, {
    error: () => undefined,
    info: () => undefined,
  });

  assert.equal(updated, false);
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
        dictionary: 'freq-dict',
        dictionaryPriority: 0,
        frequency: 77,
        displayValue: '77',
        displayValueParsed: true,
      },
      {
        term: '鍛える',
        reading: 'きたえる',
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
  assert.equal(result[0]?.frequency, 77);
  assert.equal(result[0]?.dictionaryPriority, 0);
  assert.equal(result[1]?.term, '鍛える');
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
