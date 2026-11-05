import { useEffect, useState } from 'react';
import type { GroupBy, PresetTimeRange, TimeRange } from '../../hooks/useTrends';
import { epochDayToIsoDate, isoDateToEpochDay, todayLocalDay } from '../../lib/formatters';
import type { StatsDayRange } from '../../types/stats';

interface DateRangeSelectorProps {
  range: TimeRange;
  groupBy: GroupBy;
  fillEmpty: boolean;
  /** The days currently shown, used to seed the date inputs when switching to Custom. */
  shownDays: { startDay: number; endDay: number } | null;
  onRangeChange: (r: TimeRange) => void;
  onGroupByChange: (g: GroupBy) => void;
  onFillEmptyChange: (fillEmpty: boolean) => void;
}

const CUSTOM_RANGE = 'custom';
type RangeOption = PresetTimeRange | typeof CUSTOM_RANGE;

const RANGE_OPTIONS: RangeOption[] = ['7d', '30d', '90d', '365d', 'all', CUSTOM_RANGE];
const DEFAULT_CUSTOM_DAYS = 30;
const DATE_INPUT_CLASS =
  'rounded-md border border-ctp-surface2 bg-ctp-surface0 px-1.5 py-1 text-xs text-ctp-text transition hover:border-ctp-blue';

function formatRangeOption(option: RangeOption): string {
  if (option === 'all') return 'All';
  return option === CUSTOM_RANGE ? 'Custom' : option;
}

function SegmentedControl<T extends string>({
  label,
  options,
  value,
  onChange,
  formatLabel,
}: {
  label: string;
  options: T[];
  value: T;
  onChange: (v: T) => void;
  formatLabel?: (v: T) => string;
}) {
  return (
    <div className="flex items-center gap-2">
      <span className="text-[10px] uppercase tracking-wider text-ctp-overlay1">{label}</span>
      <div className="flex bg-ctp-surface0 rounded-lg p-0.5 border border-ctp-surface1">
        {options.map((opt) => (
          <button
            key={opt}
            type="button"
            onClick={() => onChange(opt)}
            aria-pressed={value === opt}
            className={`px-2.5 py-1 rounded-md text-xs transition-colors ${
              value === opt
                ? 'bg-ctp-surface2 text-ctp-text shadow-sm'
                : 'text-ctp-overlay2 hover:text-ctp-subtext0'
            }`}
          >
            {formatLabel ? formatLabel(opt) : opt}
          </button>
        ))}
      </div>
    </div>
  );
}

interface DateDrafts {
  start: string;
  end: string;
}

/** The range two date inputs describe, or null while either is incomplete or they are reversed. */
export function parseDateDrafts(drafts: DateDrafts): StatsDayRange | null {
  const startDay = isoDateToEpochDay(drafts.start);
  const endDay = isoDateToEpochDay(drafts.end);
  return startDay !== null && endDay !== null && startDay <= endDay ? { startDay, endDay } : null;
}

/**
 * Date inputs report every keystroke, so typing 2025 passes through years 0002 and
 * 0020. Each input shows its own draft, or it would snap back and lose the digits.
 * Both drafts are checked together on each edit, so a range can move forward by
 * editing either end first. They reset whenever the range changes.
 */
function CustomRangeInputs({
  range,
  onRangeChange,
}: {
  range: StatsDayRange;
  onRangeChange: (range: StatsDayRange) => void;
}) {
  const start = epochDayToIsoDate(range.startDay);
  const end = epochDayToIsoDate(range.endDay);
  const [drafts, setDrafts] = useState<DateDrafts>({ start, end });

  useEffect(() => {
    setDrafts({ start, end });
  }, [start, end]);

  const updateDrafts = (next: DateDrafts) => {
    setDrafts(next);
    const parsed = parseDateDrafts(next);
    if (parsed) onRangeChange(parsed);
  };

  return (
    <div className="flex items-center gap-1.5 text-xs text-ctp-overlay2">
      <input
        type="date"
        aria-label="Range start"
        className={DATE_INPUT_CLASS}
        value={drafts.start}
        max={drafts.end}
        onChange={(event) => updateDrafts({ ...drafts, start: event.target.value })}
      />
      to
      <input
        type="date"
        aria-label="Range end"
        className={DATE_INPUT_CLASS}
        value={drafts.end}
        min={drafts.start}
        max={epochDayToIsoDate(todayLocalDay())}
        onChange={(event) => updateDrafts({ ...drafts, end: event.target.value })}
      />
    </div>
  );
}

export function DateRangeSelector({
  range,
  groupBy,
  fillEmpty,
  shownDays,
  onRangeChange,
  onGroupByChange,
  onFillEmptyChange,
}: DateRangeSelectorProps) {
  const selectRange = (option: RangeOption) => {
    if (option !== CUSTOM_RANGE) {
      onRangeChange(option);
      return;
    }
    if (typeof range !== 'string') return;
    const today = todayLocalDay();
    onRangeChange(shownDays ?? { startDay: today - (DEFAULT_CUSTOM_DAYS - 1), endDay: today });
  };

  return (
    <div className="flex flex-wrap items-center gap-4 text-sm">
      <SegmentedControl
        label="Range"
        options={RANGE_OPTIONS}
        value={typeof range === 'string' ? range : CUSTOM_RANGE}
        onChange={selectRange}
        formatLabel={formatRangeOption}
      />
      {typeof range === 'string' ? null : (
        <CustomRangeInputs range={range} onRangeChange={onRangeChange} />
      )}
      <SegmentedControl
        label="Group by"
        options={['day', 'week', 'month'] as GroupBy[]}
        value={groupBy}
        onChange={onGroupByChange}
        formatLabel={(g) => g.charAt(0).toUpperCase() + g.slice(1)}
      />
      <SegmentedControl
        label={`Empty ${groupBy}s`}
        options={['show', 'hide']}
        value={fillEmpty ? 'show' : 'hide'}
        onChange={(v) => onFillEmptyChange(v === 'show')}
        formatLabel={(v) => (v === 'show' ? 'Show' : 'Hide')}
      />
    </div>
  );
}
