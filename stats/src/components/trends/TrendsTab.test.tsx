import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { parseDateDrafts } from './DateRangeSelector';
import { HEATMAP_HOURS, heatmapCellOpacity } from './DayHourHeatmap';
import { describePeriodChange, SummaryStrip } from './SummaryStrip';
import { isDatalistSelection } from './TitleFilter';
import { isIsolatedPoint, startFromZero } from './TrendChart';

test('heatmap hours start at 06:00 so a late evening is not split across the axis', () => {
  assert.equal(HEATMAP_HOURS[0], 6);
  assert.equal(HEATMAP_HOURS.indexOf(23) + 1, HEATMAP_HOURS.indexOf(0));
  assert.equal(HEATMAP_HOURS.length, 24);
});

test('heatmap opacity leaves empty cells uncolored and keeps small values visible', () => {
  assert.equal(heatmapCellOpacity(0, 100), 0);
  assert.ok(heatmapCellOpacity(1, 100) > 0.2);
  assert.equal(heatmapCellOpacity(100, 100), 1);
});

for (const c of [
  { name: 'a decrease', current: 88, previous: 100, expected: '\u221212% vs previous 30 days' },
  { name: 'an increase', current: 150, previous: 100, expected: '+50% vs previous 30 days' },
  { name: 'no change', current: 100, previous: 100, expected: 'Same as previous 30 days' },
  {
    name: 'a tiny change',
    current: 1001,
    previous: 1000,
    expected: 'Under 1% from previous 30 days',
  },
  { name: 'nothing before', current: 5, previous: 0, expected: 'None in previous 30 days' },
  { name: 'nothing at all', current: 0, previous: 0, expected: null },
]) {
  test(`describePeriodChange reports ${c.name}`, () => {
    assert.equal(describePeriodChange(c.current, c.previous, 30), c.expected);
  });
}

test('SummaryStrip hides new words while filtered to a title and comparison on all-time', () => {
  const markup = renderToStaticMarkup(
    <SummaryStrip
      days={231}
      summary={{
        current: {
          watchTimeMin: 90,
          words: 24_000,
          newWords: null,
          cards: 3,
          episodes: 1,
          activeDays: 2,
        },
        previous: null,
      }}
    />,
  );
  assert.match(markup, /1\.5h/);
  assert.doesNotMatch(markup, /New words/);
  assert.doesNotMatch(markup, /previous/);
});

const SUMMARY_CURRENT = {
  watchTimeMin: 210,
  words: 15_900,
  newWords: null,
  cards: 35,
  episodes: 11,
  activeDays: 3,
};
const EMPTY_PERIOD = {
  watchTimeMin: 0,
  words: 0,
  newWords: null,
  cards: 0,
  episodes: 0,
  activeDays: 0,
};

test('SummaryStrip shows one note instead of repeating an empty previous period', () => {
  const markup = renderToStaticMarkup(
    <SummaryStrip days={30} summary={{ current: SUMMARY_CURRENT, previous: EMPTY_PERIOD }} />,
  );
  assert.equal(markup.match(/previous 30/g)?.length, 1);
  assert.match(markup, /Not watched in the previous 30 days/);
  assert.doesNotMatch(markup, /None in previous/);
});

test('SummaryStrip still compares each tile when the previous period has any activity', () => {
  const markup = renderToStaticMarkup(
    <SummaryStrip
      days={30}
      summary={{
        current: SUMMARY_CURRENT,
        previous: { ...EMPTY_PERIOD, watchTimeMin: 105, words: 7_950, activeDays: 1 },
      }}
    />,
  );
  assert.match(markup, /\+100% vs previous 30 days/);
  assert.match(markup, /None in previous 30 days/);
  assert.match(markup, /1 in previous 30/);
  assert.doesNotMatch(markup, /Not watched/);
});

test('startFromZero leads a running total with a zero point', () => {
  const points = [
    { label: 'Oct 24', value: 106 },
    { label: 'Oct 25', value: 155 },
  ];
  assert.deepEqual(startFromZero(points), [{ label: 'Start', value: 0 }, ...points]);
  assert.deepEqual(startFromZero([]), []);
});

test('isIsolatedPoint is true only for a value with gaps on both sides', () => {
  const values = [null, 3, null, 4, 5, null];
  assert.equal(isIsolatedPoint(values, 1), true);
  assert.equal(isIsolatedPoint(values, 3), false);
  assert.equal(isIsolatedPoint(values, 2), false);
});

test('isDatalistSelection separates picking a suggestion from typing', () => {
  const inputEvent = (inputType: string) => Object.assign(new Event('input'), { inputType });
  assert.equal(isDatalistSelection(new Event('input')), true);
  assert.equal(isDatalistSelection(inputEvent('insertReplacementText')), true);
  assert.equal(isDatalistSelection(inputEvent('insertText')), false);
  assert.equal(isDatalistSelection(inputEvent('deleteContentBackward')), false);
});

test('parseDateDrafts waits for two complete dates in order', () => {
  const end = '2026-03-01';
  assert.deepEqual(parseDateDrafts({ start: '2026-02-01', end }), {
    startDay: 20_485,
    endDay: 20_513,
  });
  // What a date input reports partway through typing a year, or when cleared.
  assert.equal(parseDateDrafts({ start: '0002-02-01', end }), null);
  assert.equal(parseDateDrafts({ start: '', end }), null);
  assert.equal(parseDateDrafts({ start: '2026-03-02', end }), null);
});
