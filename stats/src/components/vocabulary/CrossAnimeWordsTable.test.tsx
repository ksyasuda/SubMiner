import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { buildCrossAnimeWordRows, CrossAnimeWordsTable } from './CrossAnimeWordsTable';
import type { VocabularyEntry } from '../../types/stats';

function makeEntry(over: Partial<VocabularyEntry>): VocabularyEntry {
  return {
    wordId: 1,
    headword: '日本語',
    word: '日本語',
    reading: 'にほんご',
    frequency: 5,
    frequencyRank: 100,
    animeCount: 2,
    partOfSpeech: null,
    firstSeen: 0,
    lastSeen: 0,
    ...over,
  } as VocabularyEntry;
}

function withLocalStorage<T>(initial: Record<string, string>, run: () => T): T {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const values = new Map(Object.entries(initial));
  const storage = {
    get length() {
      return values.size;
    },
    clear() {
      values.clear();
    },
    getItem(key: string) {
      return values.get(key) ?? null;
    },
    key(index: number) {
      return Array.from(values.keys())[index] ?? null;
    },
    removeItem(key: string) {
      values.delete(key);
    },
    setItem(key: string, value: string) {
      values.set(key, value);
    },
  } as Storage;

  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: storage,
  });

  try {
    return run();
  } finally {
    if (previous) {
      Object.defineProperty(globalThis, 'localStorage', previous);
    } else {
      delete (globalThis as { localStorage?: unknown }).localStorage;
    }
  }
}

test('cross-title rows can hide kana-only headwords', () => {
  const rows = buildCrossAnimeWordRows(
    [
      makeEntry({ wordId: 1, headword: 'さらに', word: 'さらに', reading: 'さらに' }),
      makeEntry({ wordId: 2, headword: '前に', word: '前に', reading: 'まえに' }),
      makeEntry({ wordId: 3, headword: 'バカ', word: 'バカ', reading: 'バカ' }),
    ],
    new Set(),
    { hideKnown: false, hideKanaOnly: true },
  );

  assert.deepEqual(
    rows.map((row) => row.headword),
    ['前に'],
  );
});

test('cross-title table renders a Hide Kana filter button', () => {
  const markup = renderToStaticMarkup(
    <CrossAnimeWordsTable
      words={[makeEntry({ headword: 'さらに', word: 'さらに', reading: 'さらに' })]}
      knownWords={new Set()}
    />,
  );

  assert.match(markup, /Hide Kana/);
});

test('cross-title table uses saved Hide Kana preference on first render', () => {
  const markup = withLocalStorage({ 'subminer.stats.crossAnimeWords.hideKanaOnly': 'true' }, () =>
    renderToStaticMarkup(
      <CrossAnimeWordsTable
        words={[
          makeEntry({ wordId: 1, headword: 'さらに', word: 'さらに', reading: 'さらに' }),
          makeEntry({ wordId: 2, headword: '前に', word: '前に', reading: 'まえに' }),
        ]}
        knownWords={new Set()}
      />,
    ),
  );

  assert.doesNotMatch(markup, />さらに</);
  assert.match(markup, />前に</);
});
