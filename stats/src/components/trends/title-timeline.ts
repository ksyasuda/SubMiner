import type { TrendPerAnimePoint } from '../../types/stats';
import { epochDayToDate } from '../../lib/formatters';

// Past this span a daily column is too thin to see, so columns become weeks.
const MAX_DAILY_COLUMN_SPAN = 120;
const WEEK_DAYS = 7;
// Past this span a tick per month no longer fits, so ticks mark quarters.
const MAX_MONTHLY_TICK_SPAN = 400;
// The start-date tick is dropped when a month tick would sit on top of it.
const MIN_START_TICK_CLEARANCE_PERCENT = 8;

export interface TimelineCell {
  startDay: number;
  /** Position and size along the row, as percentages of its width. */
  leftPercent: number;
  widthPercent: number;
  minutes: number;
}

export interface TimelineTick {
  label: string;
  leftPercent: number;
}

export interface TitleTimelineModel {
  columnDays: number;
  ticks: TimelineTick[];
  cellsByTitle: Map<string, TimelineCell[]>;
  maxCellMinutes: number;
}

function buildTicks(startDay: number, endDay: number): TimelineTick[] {
  const span = endDay - startDay + 1;
  const monthStep = span > MAX_MONTHLY_TICK_SPAN ? 3 : 1;
  const monthTicks: TimelineTick[] = [];
  for (let day = startDay; day <= endDay; day += 1) {
    const date = epochDayToDate(day);
    if (date.getDate() !== 1 || date.getMonth() % monthStep !== 0) continue;
    monthTicks.push({
      label: date.toLocaleDateString(undefined, {
        month: 'short',
        ...(date.getMonth() === 0 ? { year: 'numeric' } : {}),
      }),
      leftPercent: ((day - startDay) / span) * 100,
    });
  }
  const firstTick = monthTicks[0];
  if (firstTick && firstTick.leftPercent < MIN_START_TICK_CLEARANCE_PERCENT) return monthTicks;
  const startTick = {
    label: epochDayToDate(startDay).toLocaleDateString(undefined, {
      month: 'short',
      day: 'numeric',
    }),
    leftPercent: 0,
  };
  return [startTick, ...monthTicks];
}

/** Lays per-title daily minutes out as positioned cells along the dashboard's date range. */
export function buildTitleTimeline(
  points: readonly TrendPerAnimePoint[],
  window: { startDay: number; endDay: number },
): TitleTimelineModel {
  const span = window.endDay - window.startDay + 1;
  const columnDays = span > MAX_DAILY_COLUMN_SPAN ? WEEK_DAYS : 1;
  const columnCount = Math.ceil(span / columnDays);
  const widthPercent = 100 / columnCount;

  const minutesByTitle = new Map<string, Map<number, number>>();
  for (const point of points) {
    const column = Math.floor((point.epochDay - window.startDay) / columnDays);
    if (column < 0 || column >= columnCount) continue;
    const columns = minutesByTitle.get(point.animeTitle) ?? new Map<number, number>();
    columns.set(column, (columns.get(column) ?? 0) + point.value);
    minutesByTitle.set(point.animeTitle, columns);
  }

  let maxCellMinutes = 0;
  const cellsByTitle = new Map<string, TimelineCell[]>();
  for (const [title, columns] of minutesByTitle) {
    const cells = [...columns.entries()]
      .sort(([left], [right]) => left - right)
      .map(([column, minutes]) => {
        maxCellMinutes = Math.max(maxCellMinutes, minutes);
        return {
          startDay: window.startDay + column * columnDays,
          leftPercent: column * widthPercent,
          widthPercent,
          minutes,
        };
      });
    cellsByTitle.set(title, cells);
  }

  return {
    columnDays,
    ticks: buildTicks(window.startDay, window.endDay),
    cellsByTitle,
    maxCellMinutes,
  };
}
