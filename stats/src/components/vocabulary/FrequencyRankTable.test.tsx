import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { FrequencyRankTable } from './FrequencyRankTable';
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
  assert.ok(!markup.includes('【カレー】'), 'should not render reading in brackets when equal to headword');
});
