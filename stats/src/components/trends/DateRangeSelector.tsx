import type { TimeRange, GroupBy } from '../../hooks/useTrends';

interface DateRangeSelectorProps {
  range: TimeRange;
  groupBy: GroupBy;
  onRangeChange: (r: TimeRange) => void;
  onGroupByChange: (g: GroupBy) => void;
}

export function DateRangeSelector({
  range,
  groupBy,
  onRangeChange,
  onGroupByChange,
}: DateRangeSelectorProps) {
  const ranges: TimeRange[] = ['7d', '30d', '90d', 'all'];
  const groups: GroupBy[] = ['day', 'month'];

  return (
    <div className="flex items-center gap-4 text-sm">
      <div className="flex items-center gap-1.5">
        <span className="text-[10px] uppercase tracking-wider text-ctp-overlay1 mr-1">Range</span>
        {ranges.map((r) => (
          <button
            key={r}
            onClick={() => onRangeChange(r)}
            aria-pressed={range === r}
            className={`px-2.5 py-1 rounded text-xs ${
              range === r
                ? 'bg-ctp-surface2 text-ctp-text'
                : 'text-ctp-overlay2 hover:text-ctp-subtext0'
            }`}
          >
            {r === 'all' ? 'All' : r}
          </button>
        ))}
      </div>
      <span className="text-ctp-surface2">{'\u00B7'}</span>
      <div className="flex items-center gap-1.5">
        <span className="text-[10px] uppercase tracking-wider text-ctp-overlay1 mr-1">Group by</span>
        {groups.map((g) => (
          <button
            key={g}
            onClick={() => onGroupByChange(g)}
            aria-pressed={groupBy === g}
            className={`px-2.5 py-1 rounded text-xs capitalize ${
              groupBy === g
                ? 'bg-ctp-surface2 text-ctp-text'
                : 'text-ctp-overlay2 hover:text-ctp-subtext0'
            }`}
          >
            {g}
          </button>
        ))}
      </div>
    </div>
  );
}
