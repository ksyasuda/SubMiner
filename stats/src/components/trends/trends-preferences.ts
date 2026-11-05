import type { GroupBy, PresetTimeRange, TimeRange } from '../../hooks/useTrends';
import {
  readBooleanPreference,
  readStringPreference,
  writeBooleanPreference,
  writeStringPreference,
} from '../../lib/preference-storage';

const RANGE_KEY = 'subminer-stats-trends-range';
const GROUP_BY_KEY = 'subminer-stats-trends-group-by';
const SHOW_EMPTY_DAYS_KEY = 'subminer-stats-trends-show-empty-days';
const RUNNING_TOTAL_METRIC_KEY = 'subminer-stats-trends-running-total-metric';

const PRESET_RANGES: readonly string[] = [
  '7d',
  '30d',
  '90d',
  '365d',
  'all',
] satisfies readonly PresetTimeRange[];

export const RUNNING_TOTAL_METRICS = [
  'watchTime',
  'words',
  'newWords',
  'cards',
  'episodes',
  'lookups',
] as const;
export type RunningTotalMetric = (typeof RUNNING_TOTAL_METRICS)[number];

function isPresetRange(value: string): value is PresetTimeRange {
  return PRESET_RANGES.includes(value);
}

/** Custom ranges are stored as `startDay:endDay`; anything unreadable falls back to 30d. */
export function loadRange(): TimeRange {
  const raw = readStringPreference(RANGE_KEY);
  if (raw === null) return '30d';
  if (isPresetRange(raw)) return raw;
  const match = /^(\d+):(\d+)$/.exec(raw);
  if (!match) return '30d';
  const startDay = Number(match[1]);
  const endDay = Number(match[2]);
  return startDay <= endDay ? { startDay, endDay } : '30d';
}

export function saveRange(range: TimeRange): void {
  writeStringPreference(
    RANGE_KEY,
    typeof range === 'string' ? range : `${range.startDay}:${range.endDay}`,
  );
}

export function loadGroupBy(): GroupBy {
  const raw = readStringPreference(GROUP_BY_KEY);
  return raw === 'week' || raw === 'month' ? raw : 'day';
}

export function saveGroupBy(groupBy: GroupBy): void {
  writeStringPreference(GROUP_BY_KEY, groupBy);
}

export function loadShowEmptyDays(): boolean {
  return readBooleanPreference(SHOW_EMPTY_DAYS_KEY, true);
}

export function saveShowEmptyDays(show: boolean): void {
  writeBooleanPreference(SHOW_EMPTY_DAYS_KEY, show);
}

export function loadRunningTotalMetric(): RunningTotalMetric {
  const raw = readStringPreference(RUNNING_TOTAL_METRIC_KEY);
  return RUNNING_TOTAL_METRICS.find((metric) => metric === raw) ?? 'watchTime';
}

export function saveRunningTotalMetric(metric: RunningTotalMetric): void {
  writeStringPreference(RUNNING_TOTAL_METRIC_KEY, metric);
}
