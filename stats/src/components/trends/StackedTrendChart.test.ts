import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildLineData,
  sortTooltipEntries,
  tooltipColumnCount,
  type PerAnimeDataPoint,
} from './StackedTrendChart';

function makePoints(titleCount: number): PerAnimeDataPoint[] {
  const points: PerAnimeDataPoint[] = [];
  for (let i = 0; i < titleCount; i += 1) {
    points.push({ epochDay: 20_000 + i, animeTitle: `Title ${i}`, value: titleCount - i });
  }
  return points;
}

test('sortTooltipEntries orders rows by value descending, tolerating string values', () => {
  const sorted = sortTooltipEntries([
    { name: 'A', value: 5 },
    { name: 'B', value: 20 },
    { name: 'C', value: '12.5' },
    { name: 'D', value: undefined },
  ]);

  assert.deepEqual(
    sorted.map((entry) => entry.name),
    ['B', 'C', 'A', 'D'],
  );
});

test('tooltipColumnCount wraps into extra columns as items grow, capped at 3', () => {
  assert.equal(tooltipColumnCount(1), 1);
  assert.equal(tooltipColumnCount(8), 1);
  assert.equal(tooltipColumnCount(9), 2);
  assert.equal(tooltipColumnCount(16), 2);
  assert.equal(tooltipColumnCount(17), 3);
  assert.equal(tooltipColumnCount(40), 3);
});

test('buildLineData keeps every title as a series instead of capping at the top 7', () => {
  const { points, seriesKeys } = buildLineData(makePoints(17));

  assert.equal(seriesKeys.length, 17);
  assert.ok(seriesKeys.includes('Title 16'));
  for (const row of points) {
    for (const key of seriesKeys) {
      assert.ok(key in row);
    }
  }
});

test('buildLineData caps series at maxSeries when set, keeping the top titles', () => {
  const { points, seriesKeys } = buildLineData(makePoints(17), 7);

  assert.equal(seriesKeys.length, 7);
  assert.deepEqual(
    seriesKeys,
    Array.from({ length: 7 }, (_, i) => `Title ${i}`),
  );
  for (const row of points) {
    assert.ok(!('Title 16' in row));
  }
});

test('buildLineData orders series by total value descending', () => {
  const { seriesKeys } = buildLineData([
    { epochDay: 20_000, animeTitle: 'Small', value: 1 },
    { epochDay: 20_000, animeTitle: 'Big', value: 10 },
    { epochDay: 20_001, animeTitle: 'Medium', value: 5 },
  ]);

  assert.deepEqual(seriesKeys, ['Big', 'Medium', 'Small']);
});

test('buildLineData recent mode keeps the most recently active titles over the largest', () => {
  // "Old Giant" has a huge cumulative total but stopped growing early; "Fresh"
  // is small but its cumulative value increased on the latest day.
  const raw: PerAnimeDataPoint[] = [
    { epochDay: 20_000, animeTitle: 'Old Giant', value: 500 },
    { epochDay: 20_001, animeTitle: 'Old Giant', value: 500 },
    { epochDay: 20_002, animeTitle: 'Old Giant', value: 500 },
    { epochDay: 20_000, animeTitle: 'Fresh', value: 1 },
    { epochDay: 20_001, animeTitle: 'Fresh', value: 2 },
    { epochDay: 20_002, animeTitle: 'Fresh', value: 5 },
  ];

  const total = buildLineData(raw, 1, 'total');
  assert.deepEqual(total.seriesKeys, ['Old Giant']);

  const recent = buildLineData(raw, 1, 'recent');
  assert.deepEqual(recent.seriesKeys, ['Fresh']);
});

test('buildLineData recent mode breaks ties by total then name', () => {
  // Both titles last increased on day 20_002; the larger total wins the tie.
  const raw: PerAnimeDataPoint[] = [
    { epochDay: 20_001, animeTitle: 'A', value: 2 },
    { epochDay: 20_002, animeTitle: 'A', value: 4 },
    { epochDay: 20_001, animeTitle: 'B', value: 3 },
    { epochDay: 20_002, animeTitle: 'B', value: 9 },
  ];

  const { seriesKeys } = buildLineData(raw, 1, 'recent');
  assert.deepEqual(seriesKeys, ['B']);
});
