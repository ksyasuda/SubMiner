import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { buildCrossAnimeWordRows, CrossAnimeWordsTable } from './CrossAnimeWordsTable';
import type { VocabularyEntry } from '../../types/stats';
import { withLocalStorage } from '../../test-utils/dom';

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
