import assert from 'node:assert/strict';
import test from 'node:test';
import { buildTitleTimeline } from './title-timeline';

const point = (epochDay: number, animeTitle: string, value: number) => ({
  epochDay,
  animeTitle,
  value,
});

test('buildTitleTimeline positions one column per day for short ranges', () => {
  const model = buildTitleTimeline([point(20_002, 'A', 30), point(20_000, 'A', 10)], {
    startDay: 20_000,
    endDay: 20_009,
  });

  assert.equal(model.columnDays, 1);
  assert.deepEqual(
    model.cellsByTitle.get('A')?.map((cell) => [cell.startDay, cell.leftPercent, cell.minutes]),
    [
      [20_000, 0, 10],
      [20_002, 20, 30],
    ],
  );
  assert.equal(model.maxCellMinutes, 30);
});

test('buildTitleTimeline merges days into weekly columns for long ranges', () => {
  const model = buildTitleTimeline(
    [point(20_000, 'A', 10), point(20_006, 'A', 5), point(20_007, 'A', 1)],
    { startDay: 20_000, endDay: 20_199 },
  );

  assert.equal(model.columnDays, 7);
  assert.deepEqual(
    model.cellsByTitle.get('A')?.map((cell) => [cell.startDay, cell.minutes]),
    [
      [20_000, 15],
      [20_007, 1],
    ],
  );
});

test('buildTitleTimeline ignores points outside the range', () => {
  const model = buildTitleTimeline([point(19_999, 'A', 10), point(20_010, 'B', 10)], {
    startDay: 20_000,
    endDay: 20_009,
  });
  assert.equal(model.cellsByTitle.size, 0);
});

test('buildTitleTimeline labels the start of the range unless a month tick is already there', () => {
  // 2024-10-05 to 2024-11-03: only November starts inside the range.
  const midMonth = buildTitleTimeline([], { startDay: 20_001, endDay: 20_030 });
  assert.equal(midMonth.ticks.length, 2);
  assert.equal(midMonth.ticks[0]?.leftPercent, 0);

  // 2024-10-01 onward: October's own tick already marks the start.
  const monthStart = buildTitleTimeline([], { startDay: 19_997, endDay: 20_030 });
  assert.equal(monthStart.ticks.length, 2);
  assert.equal(monthStart.ticks[0]?.leftPercent, 0);
  assert.notEqual(monthStart.ticks[1]?.leftPercent, 0);
});
