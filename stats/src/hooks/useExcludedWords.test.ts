import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isExcludedWord,
  getExcludedWordsSnapshot,
  initializeExcludedWordsStore,
  resetExcludedWordsStoreForTests,
  setExcludedWords,
  subscribeExcludedWordsServerSync,
} from './useExcludedWords';
import { BASE_URL } from '../lib/api-client';

const STORAGE_KEY = 'subminer-excluded-words';

function installLocalStorage(initial: Record<string, string> = {}) {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const values = new Map(Object.entries(initial));
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
      removeItem: (key: string) => values.delete(key),
    },
  });
  return {
    values,
    restore: () => {
      if (previous) {
        Object.defineProperty(globalThis, 'localStorage', previous);
      } else {
        delete (globalThis as { localStorage?: unknown }).localStorage;
      }
    },
  };
}

test('initializeExcludedWordsStore seeds empty database exclusions from localStorage', async () => {
  resetExcludedWordsStoreForTests();
  const localRows = [{ headword: '猫', word: '猫', reading: 'ねこ' }];
  const { values: storage, restore } = installLocalStorage({
    [STORAGE_KEY]: JSON.stringify(localRows),
  });
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
    restore();
    resetExcludedWordsStoreForTests();
  }
});

test('setExcludedWords updates the database-backed exclusion list', async () => {
  resetExcludedWordsStoreForTests();
  const { values: storage, restore } = installLocalStorage();
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
    restore();
    resetExcludedWordsStoreForTests();
  }
});

test('setExcludedWords persists one row per excluded token', async () => {
  resetExcludedWordsStoreForTests();
  const { values: storage, restore } = installLocalStorage();
  const originalFetch = globalThis.fetch;
  let seenBody = '';
  globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    seenBody = String(init?.body ?? '');
    return new Response(JSON.stringify({ ok: true }), { status: 200 });
  }) as typeof globalThis.fetch;

  try {
    const rows = [
      { headword: 'ない', word: 'ない', reading: 'ない' },
      { headword: 'ない', word: '無い', reading: 'ない' },
    ];
    const expected = [{ headword: 'ない', word: 'ない', reading: 'ない' }];

    await setExcludedWords(rows);

    assert.deepEqual(getExcludedWordsSnapshot(), expected);
    assert.equal(seenBody, JSON.stringify({ words: expected }));
    assert.equal(storage.get(STORAGE_KEY), JSON.stringify(expected));
  } finally {
    globalThis.fetch = originalFetch;
    restore();
    resetExcludedWordsStoreForTests();
  }
});

test('exclusion matching covers vocabulary rows with the same visible token', () => {
  const excluded = [{ headword: 'ない', word: 'ない', reading: 'ない' }];

  assert.equal(isExcludedWord(excluded, { headword: 'ない', word: '無い', reading: 'ない' }), true);
  assert.equal(isExcludedWord(excluded, { headword: '無い', word: 'ない', reading: 'ない' }), true);
  assert.equal(
    isExcludedWord(excluded, { headword: 'なる', word: 'なる', reading: 'なる' }),
    false,
  );
});

test('setExcludedWords rolls back local state when persistence fails', async () => {
  resetExcludedWordsStoreForTests();
  const previousRows = [{ headword: '猫', word: '猫', reading: 'ねこ' }];
  const nextRows = [{ headword: 'する', word: 'する', reading: 'する' }];
  const { values: storage, restore } = installLocalStorage({
    [STORAGE_KEY]: JSON.stringify(previousRows),
  });
  const originalFetch = globalThis.fetch;
  const originalConsoleError = console.error;
  console.error = () => {};
  globalThis.fetch = (async () => {
    return new Response('failed', { status: 500 });
  }) as typeof globalThis.fetch;

  try {
    await assert.rejects(() => setExcludedWords(nextRows), /Stats API error: 500/);

    assert.deepEqual(getExcludedWordsSnapshot(), previousRows);
    assert.equal(storage.get(STORAGE_KEY), JSON.stringify(previousRows));
  } finally {
    globalThis.fetch = originalFetch;
    console.error = originalConsoleError;
    restore();
    resetExcludedWordsStoreForTests();
  }
});

test('initializeExcludedWordsStore retries after transient database load failures', async () => {
  resetExcludedWordsStoreForTests();
  const { restore } = installLocalStorage();
  const originalFetch = globalThis.fetch;
  const originalConsoleError = console.error;
  console.error = () => {};
  let calls = 0;
  globalThis.fetch = (async () => {
    calls += 1;
    if (calls === 1) {
      return new Response('failed', { status: 500 });
    }
    return new Response(JSON.stringify([{ headword: '猫', word: '猫', reading: 'ねこ' }]), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as typeof globalThis.fetch;

  try {
    await initializeExcludedWordsStore();
    await initializeExcludedWordsStore();

    assert.equal(calls, 2);
    assert.deepEqual(getExcludedWordsSnapshot(), [{ headword: '猫', word: '猫', reading: 'ねこ' }]);
  } finally {
    globalThis.fetch = originalFetch;
    console.error = originalConsoleError;
    restore();
    resetExcludedWordsStoreForTests();
  }
});

test('a failing server-sync listener neither rolls back the write nor blocks other listeners', async () => {
  resetExcludedWordsStoreForTests();
  const { values: storage, restore } = installLocalStorage();
  const originalFetch = globalThis.fetch;
  const originalConsoleError = console.error;
  console.error = () => {};
  globalThis.fetch = (async () =>
    new Response(JSON.stringify({ ok: true }), { status: 200 })) as typeof globalThis.fetch;
  const notified: string[] = [];
  const unsubscribeFirst = subscribeExcludedWordsServerSync(() => {
    notified.push('first');
    throw new Error('listener exploded');
  });
  const unsubscribeSecond = subscribeExcludedWordsServerSync(() => {
    notified.push('second');
  });

  try {
    const rows = [{ headword: 'する', word: 'する', reading: 'する' }];
    await assert.doesNotReject(() => setExcludedWords(rows));

    assert.deepEqual(notified, ['first', 'second']);
    assert.deepEqual(getExcludedWordsSnapshot(), rows);
    assert.equal(storage.get(STORAGE_KEY), JSON.stringify(rows));
  } finally {
    unsubscribeFirst();
    unsubscribeSecond();
    globalThis.fetch = originalFetch;
    console.error = originalConsoleError;
    restore();
    resetExcludedWordsStoreForTests();
  }
});

test('overlapping writes serialize so an older list cannot overwrite a newer edit', async () => {
  resetExcludedWordsStoreForTests();
  const { restore } = installLocalStorage();
  const originalFetch = globalThis.fetch;
  const sentBodies: string[] = [];
  let releaseFirst: (() => void) | null = null;
  globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    sentBodies.push(String(init?.body ?? ''));
    if (sentBodies.length === 1) {
      await new Promise<void>((resolve) => {
        releaseFirst = resolve;
      });
    }
    return new Response(JSON.stringify({ ok: true }), { status: 200 });
  }) as typeof globalThis.fetch;
  const syncs: string[] = [];
  const unsubscribe = subscribeExcludedWordsServerSync(() => {
    syncs.push(JSON.stringify(getExcludedWordsSnapshot()));
  });

  try {
    const first = [{ headword: '猫', word: '猫', reading: 'ねこ' }];
    const second = [...first, { headword: '犬', word: '犬', reading: 'いぬ' }];
    const third = [...second, { headword: '鳥', word: '鳥', reading: 'とり' }];

    // The first write reaches the network before the later edits are made.
    const firstWrite = setExcludedWords(first);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const secondWrite = setExcludedWords(second);
    const thirdWrite = setExcludedWords(third);

    const release = releaseFirst as (() => void) | null;
    assert.ok(release, 'expected the first write to be in flight');
    release();
    await Promise.all([firstWrite, secondWrite, thirdWrite]);

    // The in-flight write finishes first, the superseded middle write is
    // dropped, and the newest list is the last thing the server is told.
    assert.deepEqual(sentBodies, [
      JSON.stringify({ words: first }),
      JSON.stringify({ words: third }),
    ]);
    assert.deepEqual(getExcludedWordsSnapshot(), third);
    assert.equal(syncs.length, 2);
    assert.equal(syncs.at(-1), JSON.stringify(third));
  } finally {
    unsubscribe();
    globalThis.fetch = originalFetch;
    restore();
    resetExcludedWordsStoreForTests();
  }
});
