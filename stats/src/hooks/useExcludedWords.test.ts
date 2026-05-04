import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getExcludedWordsSnapshot,
  initializeExcludedWordsStore,
  resetExcludedWordsStoreForTests,
  setExcludedWords,
} from './useExcludedWords';
import { BASE_URL } from '../lib/api-client';

const STORAGE_KEY = 'subminer-excluded-words';

function installLocalStorage(initial: Record<string, string> = {}) {
  const values = new Map(Object.entries(initial));
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
      removeItem: (key: string) => values.delete(key),
    },
  });
  return values;
}

test('initializeExcludedWordsStore seeds empty database exclusions from localStorage', async () => {
  resetExcludedWordsStoreForTests();
  const localRows = [{ headword: '猫', word: '猫', reading: 'ねこ' }];
  const storage = installLocalStorage({ [STORAGE_KEY]: JSON.stringify(localRows) });
  const originalFetch = globalThis.fetch;
  const requests: Array<{ url: string; method: string; body: string }> = [];
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    requests.push({
      url: String(input),
      method: init?.method ?? 'GET',
      body: String(init?.body ?? ''),
    });
    if (!init?.method) {
      return new Response(JSON.stringify([]), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      });
    }
    return new Response(JSON.stringify({ ok: true }), { status: 200 });
  }) as typeof globalThis.fetch;

  try {
    await initializeExcludedWordsStore();

    assert.deepEqual(getExcludedWordsSnapshot(), localRows);
    assert.deepEqual(requests, [
      { url: `${BASE_URL}/api/stats/excluded-words`, method: 'GET', body: '' },
      {
        url: `${BASE_URL}/api/stats/excluded-words`,
        method: 'PUT',
        body: JSON.stringify({ words: localRows }),
      },
    ]);
    assert.equal(storage.get(STORAGE_KEY), JSON.stringify(localRows));
  } finally {
    globalThis.fetch = originalFetch;
    resetExcludedWordsStoreForTests();
  }
});

test('setExcludedWords updates the database-backed exclusion list', async () => {
  resetExcludedWordsStoreForTests();
  const storage = installLocalStorage();
  const originalFetch = globalThis.fetch;
  let seenBody = '';
  globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    seenBody = String(init?.body ?? '');
    return new Response(JSON.stringify({ ok: true }), { status: 200 });
  }) as typeof globalThis.fetch;

  try {
    const rows = [{ headword: 'する', word: 'する', reading: 'する' }];
    await setExcludedWords(rows);

    assert.deepEqual(getExcludedWordsSnapshot(), rows);
    assert.equal(seenBody, JSON.stringify({ words: rows }));
    assert.equal(storage.get(STORAGE_KEY), JSON.stringify(rows));
  } finally {
    globalThis.fetch = originalFetch;
    resetExcludedWordsStoreForTests();
  }
});
