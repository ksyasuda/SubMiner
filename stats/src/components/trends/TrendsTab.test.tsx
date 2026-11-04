import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { HEATMAP_HOURS, heatmapCellOpacity } from './DayHourHeatmap';
import { describePeriodChange, SummaryStrip } from './SummaryStrip';
import { isIsolatedPoint } from './TrendChart';

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

test('isIsolatedPoint is true only for a value with gaps on both sides', () => {
  const values = [null, 3, null, 4, 5, null];
  assert.equal(isIsolatedPoint(values, 1), true);
  assert.equal(isIsolatedPoint(values, 3), false);
  assert.equal(isIsolatedPoint(values, 2), false);
});
