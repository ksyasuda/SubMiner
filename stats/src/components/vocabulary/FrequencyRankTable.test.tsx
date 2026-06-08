import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import {
  buildFrequencyRankRows,
  FrequencyRankTable,
  isKanaOnlyTokenText,
} from './FrequencyRankTable';
import type { VocabularyEntry } from '../../types/stats';

function makeEntry(over: Partial<VocabularyEntry>): VocabularyEntry {
  return {
    wordId: 1,
    headword: '日本語',
    word: '日本語',
    reading: 'にほんご',
    frequency: 5,
    frequencyRank: 100,
    animeCount: 1,
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

test('renders headword and reading inline in a single column (no separate Reading header)', () => {
  const entry = makeEntry({});
  const markup = renderToStaticMarkup(
    <FrequencyRankTable words={[entry]} knownWords={new Set()} />,
  );
  assert.ok(!markup.includes('>Reading<'), 'should not have a Reading column header');
  assert.ok(markup.includes('日本語'), 'should include the headword');
  assert.ok(markup.includes('にほんご'), 'should include the reading inline');
});

test('omits reading when reading equals headword', () => {
  const entry = makeEntry({ headword: 'カレー', word: 'カレー', reading: 'カレー' });
  const markup = renderToStaticMarkup(
    <FrequencyRankTable words={[entry]} knownWords={new Set()} />,
  );
  assert.ok(markup.includes('カレー'), 'should include the headword');
  assert.ok(
    !markup.includes('【'),
    'should not render any bracketed reading when equal to headword',
  );
});

test('identifies kana-only token text without hiding mixed kanji words', () => {
  assert.equal(isKanaOnlyTokenText('さらに'), true);
  assert.equal(isKanaOnlyTokenText('バカ'), true);
  assert.equal(isKanaOnlyTokenText('カレー'), true);
  assert.equal(isKanaOnlyTokenText('前に'), false);
  assert.equal(isKanaOnlyTokenText('間違いない'), false);
});

test('frequency rows can hide kana-only headwords', () => {
  const rows = buildFrequencyRankRows(
    [
      makeEntry({ wordId: 1, headword: 'さらに', word: 'さらに', frequencyRank: 10 }),
      makeEntry({
        wordId: 2,
        headword: '前に',
        word: '前に',
        reading: 'まえに',
        frequencyRank: 20,
      }),
      makeEntry({ wordId: 3, headword: 'バカ', word: 'バカ', reading: 'バカ', frequencyRank: 30 }),
    ],
    new Set(),
    { hideKnown: false, hideKanaOnly: true },
  );

  assert.deepEqual(
    rows.map((row) => row.headword),
    ['前に'],
  );
});

test('renders a Hide Kana filter button', () => {
  const entry = makeEntry({ headword: 'さらに', word: 'さらに', reading: 'さらに' });
  const markup = renderToStaticMarkup(
    <FrequencyRankTable words={[entry]} knownWords={new Set()} />,
  );
  assert.match(markup, /Hide Kana/);
});

test('uses saved Hide Kana preference on first render', () => {
  const markup = withLocalStorage({ 'subminer.stats.frequencyRank.hideKanaOnly': 'true' }, () =>
    renderToStaticMarkup(
      <FrequencyRankTable
        words={[
          makeEntry({ wordId: 1, headword: 'さらに', word: 'さらに', frequencyRank: 10 }),
          makeEntry({
            wordId: 2,
            headword: '前に',
            word: '前に',
            reading: 'まえに',
            frequencyRank: 20,
          }),
        ]}
        knownWords={new Set()}
      />,
    ),
  );

  assert.doesNotMatch(markup, />さらに</);
  assert.match(markup, />前に</);
});

test('uses saved Hide Known preference on first render', () => {
  const markup = withLocalStorage({ 'subminer.stats.frequencyRank.hideKnown': 'false' }, () =>
    renderToStaticMarkup(
      <FrequencyRankTable
        words={[makeEntry({ headword: '日本語', word: '日本語', frequencyRank: 10 })]}
        knownWords={new Set(['日本語'])}
      />,
    ),
  );

  assert.match(markup, />Most Common Words Seen</);
  assert.match(markup, />日本語</);
});
