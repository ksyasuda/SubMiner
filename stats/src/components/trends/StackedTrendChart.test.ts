import assert from 'node:assert/strict';
import test from 'node:test';
import { buildLineData, type PerAnimeDataPoint } from './StackedTrendChart';

function makePoints(titleCount: number): PerAnimeDataPoint[] {
  const points: PerAnimeDataPoint[] = [];
  for (let i = 0; i < titleCount; i += 1) {
    points.push({ epochDay: 20_000 + i, animeTitle: `Title ${i}`, value: titleCount - i });
  }
  return points;
}

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
