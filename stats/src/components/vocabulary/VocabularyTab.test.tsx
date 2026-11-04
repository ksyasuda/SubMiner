import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { KanjiBreakdown, kanjiCellAlpha, kanjiInKnownWords } from './KanjiBreakdown';
import { VocabularyCoverage } from './VocabularyCoverage';
import { highlightWord } from './WordDetailPanel';
import { buildVocabularyGrowth } from './VocabularyTab';
import {
  buildWordQuery,
  defaultSortForView,
  nextSort,
  type WordExplorerState,
} from './word-explorer-query';

const baseState: WordExplorerState = {
  view: 'all',
  search: '',
  hideKnown: false,
  hideNames: true,
  hideKana: true,
  partOfSpeech: null,
  sort: defaultSortForView('all'),
  page: 2,
};

test('Worth mining means common and unknown, whatever the Hide known chip says', () => {
  const query = buildWordQuery({ ...baseState, view: 'worth', sort: defaultSortForView('worth') });
  assert.equal(query.hideKnown, true);
  assert.equal(query.maxRank, 10_000);
  assert.equal(query.sort, 'seen');
  assert.equal(query.offset, 50);
});

test('Across titles needs two titles and sorts by them', () => {
  const query = buildWordQuery({
    ...baseState,
    view: 'titles',
    sort: defaultSortForView('titles'),
  });
  assert.equal(query.minTitles, 2);
  assert.equal(query.sort, 'titles');
});

test('searching ignores the view and filters so any word can be found', () => {
  const query = buildWordQuery({
    ...baseState,
    view: 'worth',
    search: ' 猫 ',
    hideKnown: true,
    partOfSpeech: 'verb',
  });
  assert.equal(query.search, '猫');
  assert.equal(query.hideKnown, false);
  assert.equal(query.hideNames, false);
  assert.equal(query.hideKana, false);
  assert.equal(query.partOfSpeech, undefined);
  assert.equal(query.maxRank, undefined);
});

test('nextSort flips the active column and starts new ones on their useful end', () => {
  assert.deepEqual(nextSort({ sort: 'seen', descending: true }, 'seen'), {
    sort: 'seen',
    descending: false,
  });
  assert.deepEqual(nextSort({ sort: 'seen', descending: true }, 'rank'), {
    sort: 'rank',
    descending: false,
  });
  assert.deepEqual(nextSort({ sort: 'rank', descending: false }, 'titles'), {
    sort: 'titles',
    descending: true,
  });
});

test('vocabulary growth is a running total', () => {
  const growth = buildVocabularyGrowth([
    { epochDay: 20_000, wordCount: 5 },
    { epochDay: 20_010, wordCount: 2 },
  ]);
  assert.deepEqual(
    growth.map((point) => point.value),
    [5, 7],
  );
});

test('kanjiInKnownWords collects kanji and skips kana and Latin', () => {
  assert.deepEqual(
    [...kanjiInKnownWords(['見る', 'さん (used with adults)', '大丈夫'])].sort(),
    ['丈', '夫', '大', '見'].sort(),
  );
});

test('kanjiCellAlpha keeps rare kanji visible and scales up to the most seen', () => {
  assert.ok(kanjiCellAlpha(1, 7_000) > 0.08);
  assert.ok(kanjiCellAlpha(900, 7_000) < kanjiCellAlpha(4_000, 7_000));
  assert.equal(kanjiCellAlpha(7_000, 7_000), 0.6);
});

const KANJI = [
  { kanjiId: 1, kanji: '人', frequency: 100, firstSeen: 0, lastSeen: 0 },
  { kanjiId: 2, kanji: '鈴', frequency: 50, firstSeen: 0, lastSeen: 0 },
];

test('KanjiBreakdown reports the real totals and flags kanji outside known words', () => {
  const markup = renderToStaticMarkup(
    <KanjiBreakdown kanji={KANJI} knownWords={new Set(['人間'])} />,
  );
  assert.match(markup, /All 2/);
  assert.match(markup, /Not in a known word 1/);
  assert.equal(markup.match(/outline-dashed/g)?.length, 1);
  assert.match(markup, /lang="ja"/);
});

test('KanjiBreakdown drops known-word features when there is no known-word data', () => {
  const markup = renderToStaticMarkup(<KanjiBreakdown kanji={KANJI} knownWords={new Set()} />);
  assert.doesNotMatch(markup, /Not in a known word/);
  assert.doesNotMatch(markup, /outline-dashed/);
});

test('VocabularyCoverage shows each band as a share of the words seen in it', () => {
  const markup = renderToStaticMarkup(
    <VocabularyCoverage
      bands={[
        { band: 'top1k', seen: 910, known: 690 },
        { band: 'top5k', seen: 0, known: 0 },
      ]}
    />,
  );
  assert.match(markup, /690 of 910, 76%/);
  assert.match(markup, /0 of 0, 0%/);
});

test('highlightWord marks the longest form, whatever order forms arrive in', () => {
  const markup = renderToStaticMarkup(<p>{highlightWord('食べたい', ['食べ', '食べた'])}</p>);
  assert.match(markup, />食べた<\/mark>い/);
});
