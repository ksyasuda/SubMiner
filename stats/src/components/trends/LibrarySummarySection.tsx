import { useMemo, useState } from 'react';
import {
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import type { LibrarySummaryRow } from '../../types/stats';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from '../../lib/chart-theme';
import { epochDayToDate, formatDuration, formatNumber } from '../../lib/formatters';

interface LibrarySummarySectionProps {
  rows: LibrarySummaryRow[];
  hiddenTitles: ReadonlySet<string>;
}

const LEADERBOARD_LIMIT = 10;
const LEADERBOARD_HEIGHT = 260;
const LEADERBOARD_BAR_COLOR = '#8aadf4';
const TABLE_MAX_HEIGHT = 480;

type SortColumn =
  | 'title'
  | 'watchTimeMin'
  | 'videos'
  | 'sessions'
  | 'cards'
  | 'words'
  | 'lookups'
  | 'lookupsPerHundred'
  | 'firstWatched';

type SortDirection = 'asc' | 'desc';

interface ColumnDef {
  id: SortColumn;
  label: string;
  align: 'left' | 'right';
}

const COLUMNS: ColumnDef[] = [
  { id: 'title', label: 'Title', align: 'left' },
  { id: 'watchTimeMin', label: 'Watch Time', align: 'right' },
  { id: 'videos', label: 'Videos', align: 'right' },
  { id: 'sessions', label: 'Sessions', align: 'right' },
  { id: 'cards', label: 'Cards', align: 'right' },
  { id: 'words', label: 'Words', align: 'right' },
  { id: 'lookups', label: 'Lookups', align: 'right' },
  { id: 'lookupsPerHundred', label: 'Lookups/100w', align: 'right' },
  { id: 'firstWatched', label: 'Date Range', align: 'right' },
];

function truncateTitle(title: string, maxChars: number): string {
  if (title.length <= maxChars) return title;
  return `${title.slice(0, maxChars - 1)}…`;
}

function formatDateRange(firstEpochDay: number, lastEpochDay: number): string {
  const fmt = (epochDay: number) =>
    epochDayToDate(epochDay).toLocaleDateString(undefined, {
      month: 'short',
      day: 'numeric',
    });
  if (firstEpochDay === lastEpochDay) return fmt(firstEpochDay);
  return `${fmt(firstEpochDay)} → ${fmt(lastEpochDay)}`;
}

function formatWatchTime(min: number): string {
  return formatDuration(min * 60_000);
}

function compareRows(
  a: LibrarySummaryRow,
  b: LibrarySummaryRow,
  column: SortColumn,
  direction: SortDirection,
): number {
  const sign = direction === 'asc' ? 1 : -1;

  if (column === 'title') {
    return a.title.localeCompare(b.title) * sign;
  }

  if (column === 'firstWatched') {
    return (a.firstWatched - b.firstWatched) * sign;
  }

  if (column === 'lookupsPerHundred') {
    const aVal = a.lookupsPerHundred;
    const bVal = b.lookupsPerHundred;
    if (aVal === null && bVal === null) return 0;
    if (aVal === null) return 1;
    if (bVal === null) return -1;
    return (aVal - bVal) * sign;
  }

  const aVal = a[column] as number;
  const bVal = b[column] as number;
  return (aVal - bVal) * sign;
}

export function LibrarySummarySection({ rows, hiddenTitles }: LibrarySummarySectionProps) {
  const [sortColumn, setSortColumn] = useState<SortColumn>('watchTimeMin');
  const [sortDirection, setSortDirection] = useState<SortDirection>('desc');

  const visibleRows = useMemo(
    () => rows.filter((row) => !hiddenTitles.has(row.title)),
    [rows, hiddenTitles],
  );

  const sortedRows = useMemo(
    () => [...visibleRows].sort((a, b) => compareRows(a, b, sortColumn, sortDirection)),
    [visibleRows, sortColumn, sortDirection],
  );

  const leaderboard = useMemo(
    () =>
      [...visibleRows]
        .sort((a, b) => b.watchTimeMin - a.watchTimeMin)
        .slice(0, LEADERBOARD_LIMIT)
        .map((row) => ({
          title: row.title,
          displayTitle: truncateTitle(row.title, 24),
          watchTimeMin: row.watchTimeMin,
        })),
    [visibleRows],
  );

  if (visibleRows.length === 0) {
    return (
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <div className="text-xs text-ctp-overlay2">
          No library activity in the selected window.
        </div>
      </div>
    );
  }

  const handleHeaderClick = (column: SortColumn) => {
    if (column === sortColumn) {
      setSortDirection((prev) => (prev === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortColumn(column);
      setSortDirection(column === 'title' ? 'asc' : 'desc');
    }
  };

  return (
    <>
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <h3 className="text-xs font-semibold text-ctp-text mb-2">
          Top Titles by Watch Time (min)
        </h3>
        <ResponsiveContainer width="100%" height={LEADERBOARD_HEIGHT}>
          <BarChart
            data={leaderboard}
            layout="vertical"
            margin={{ top: 8, right: 16, bottom: 8, left: 8 }}
          >
            <CartesianGrid stroke={CHART_THEME.grid} {...CHART_DEFAULTS.grid} />
            <XAxis
              type="number"
              tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
              axisLine={{ stroke: CHART_THEME.axisLine }}
              tickLine={false}
            />
            <YAxis
              type="category"
              dataKey="displayTitle"
              width={160}
              tick={{ fontSize: CHART_DEFAULTS.tickFontSize, fill: CHART_THEME.tick }}
              axisLine={{ stroke: CHART_THEME.axisLine }}
              tickLine={false}
              interval={0}
            />
            <Tooltip
              contentStyle={TOOLTIP_CONTENT_STYLE}
              formatter={(value: number) => [`${value} min`, 'Watch Time']}
              labelFormatter={(_label, payload) => {
                const datum = payload?.[0]?.payload as { title?: string } | undefined;
                return datum?.title ?? '';
              }}
            />
            <Bar dataKey="watchTimeMin" fill={LEADERBOARD_BAR_COLOR} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <h3 className="text-xs font-semibold text-ctp-text mb-2">Per-Title Summary</h3>
        <div
          className="overflow-auto"
          style={{ maxHeight: TABLE_MAX_HEIGHT }}
        >
          <table className="w-full text-xs">
            <thead className="sticky top-0 bg-ctp-surface0">
              <tr className="border-b border-ctp-surface1 text-ctp-subtext0">
                {COLUMNS.map((column) => {
                  const isActive = column.id === sortColumn;
                  const indicator = isActive ? (sortDirection === 'asc' ? ' ▲' : ' ▼') : '';
                  return (
                    <th
                      key={column.id}
                      scope="col"
                      className={`px-2 py-2 font-medium select-none cursor-pointer hover:text-ctp-text ${
                        column.align === 'right' ? 'text-right' : 'text-left'
                      } ${isActive ? 'text-ctp-text' : ''}`}
                      onClick={() => handleHeaderClick(column.id)}
                    >
                      {column.label}
                      {indicator}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {sortedRows.map((row) => (
                <tr
                  key={row.title}
                  className="border-b border-ctp-surface1 last:border-b-0 hover:bg-ctp-surface1/40"
                >
                  <td
                    className="px-2 py-2 text-left text-ctp-text max-w-[240px] truncate"
                    title={row.title}
                  >
                    {row.title}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatWatchTime(row.watchTimeMin)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.videos)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.sessions)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.cards)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.words)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.lookups)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {row.lookupsPerHundred === null
                      ? '—'
                      : row.lookupsPerHundred.toFixed(1)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-subtext0 tabular-nums">
                    {formatDateRange(row.firstWatched, row.lastWatched)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}
