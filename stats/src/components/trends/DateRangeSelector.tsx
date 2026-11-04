import type { GroupBy, PresetTimeRange, TimeRange } from '../../hooks/useTrends';
import { epochDayToIsoDate, isoDateToEpochDay, todayLocalDay } from '../../lib/formatters';

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
        <div className="flex items-center gap-1.5 text-xs text-ctp-overlay2">
          <input
            type="date"
            aria-label="Range start"
            className={DATE_INPUT_CLASS}
            value={epochDayToIsoDate(range.startDay)}
            max={epochDayToIsoDate(range.endDay)}
            onChange={(event) => {
              const startDay = isoDateToEpochDay(event.target.value);
              if (startDay !== null && startDay <= range.endDay) {
                onRangeChange({ startDay, endDay: range.endDay });
              }
            }}
          />
          to
          <input
            type="date"
            aria-label="Range end"
            className={DATE_INPUT_CLASS}
            value={epochDayToIsoDate(range.endDay)}
            min={epochDayToIsoDate(range.startDay)}
            max={epochDayToIsoDate(todayLocalDay())}
            onChange={(event) => {
              const endDay = isoDateToEpochDay(event.target.value);
              if (endDay !== null && endDay >= range.startDay) {
                onRangeChange({ startDay: range.startDay, endDay });
              }
            }}
          />
        </div>
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
