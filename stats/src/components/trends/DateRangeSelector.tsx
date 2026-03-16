import type { TimeRange, GroupBy } from '../../hooks/useTrends';

interface DateRangeSelectorProps {
  range: TimeRange;
  groupBy: GroupBy;
  onRangeChange: (r: TimeRange) => void;
  onGroupByChange: (g: GroupBy) => void;
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
  onRangeChange,
  onGroupByChange,
}: DateRangeSelectorProps) {
  return (
    <div className="flex items-center gap-4 text-sm">
      <SegmentedControl
        label="Range"
        options={['7d', '30d', '90d', 'all'] as TimeRange[]}
        value={range}
        onChange={onRangeChange}
        formatLabel={(r) => r === 'all' ? 'All' : r}
      />
      <SegmentedControl
        label="Group by"
        options={['day', 'month'] as GroupBy[]}
        value={groupBy}
        onChange={onGroupByChange}
        formatLabel={(g) => g.charAt(0).toUpperCase() + g.slice(1)}
      />
    </div>
  );
}
