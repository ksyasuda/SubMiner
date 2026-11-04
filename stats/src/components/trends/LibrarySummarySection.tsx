import { useMemo, useState } from 'react';
import type { LibrarySummaryRow } from '../../types/stats';
import { epochDayToDate, formatDuration, formatNumber } from '../../lib/formatters';

interface LibrarySummarySectionProps {
  rows: LibrarySummaryRow[];
  selectedTitle: string | null;
  onSelectTitle: (title: string) => void;
  onOpenTitle: (row: LibrarySummaryRow) => void;
}

const TABLE_MAX_HEIGHT = 480;

type SortColumn =
  | 'title'
  | 'watchTimeMin'
  | 'videos'
  | 'episodesFinished'
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
  { id: 'episodesFinished', label: 'Finished', align: 'right' },
  { id: 'sessions', label: 'Sessions', align: 'right' },
  { id: 'cards', label: 'Cards', align: 'right' },
  { id: 'words', label: 'Words', align: 'right' },
  { id: 'lookups', label: 'Lookups', align: 'right' },
  { id: 'lookupsPerHundred', label: 'Lookups/100w', align: 'right' },
  { id: 'firstWatched', label: 'Date Range', align: 'right' },
];

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

  if (column === 'lookupsPerHundred') {
    const aVal = a.lookupsPerHundred;
    const bVal = b.lookupsPerHundred;
    if (aVal === null && bVal === null) return 0;
    if (aVal === null) return 1;
    if (bVal === null) return -1;
    return (aVal - bVal) * sign;
  }

  return (a[column] - b[column]) * sign;
}

export function LibrarySummarySection({
  rows,
  selectedTitle,
  onSelectTitle,
  onOpenTitle,
}: LibrarySummarySectionProps) {
  const [sortColumn, setSortColumn] = useState<SortColumn>('watchTimeMin');
  const [sortDirection, setSortDirection] = useState<SortDirection>('desc');

  const sortedRows = useMemo(
    () => [...rows].sort((a, b) => compareRows(a, b, sortColumn, sortDirection)),
    [rows, sortColumn, sortDirection],
  );
  const maxWatchTime = useMemo(
    () => rows.reduce((max, row) => Math.max(max, row.watchTimeMin), 0),
    [rows],
  );

  if (rows.length === 0) {
    return (
      <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
        <div className="text-xs text-ctp-overlay2">No library activity in the selected window.</div>
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
    <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
      <h3 className="text-xs font-semibold text-ctp-text mb-2">Per-Title Summary</h3>
      <div className="overflow-auto" style={{ maxHeight: TABLE_MAX_HEIGHT }}>
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
                    aria-sort={
                      isActive ? (sortDirection === 'asc' ? 'ascending' : 'descending') : undefined
                    }
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
            {sortedRows.map((row) => {
              const isSelected = row.title === selectedTitle;
              return (
                <tr
                  key={row.title}
                  className={`border-b border-ctp-surface1 last:border-b-0 hover:bg-ctp-surface1/40 ${
                    isSelected ? 'bg-ctp-surface1/60' : ''
                  }`}
                >
                  <td className="px-2 py-2 text-left max-w-[240px]">
                    <div className="flex min-w-0 items-center gap-1.5">
                      <button
                        type="button"
                        aria-pressed={isSelected}
                        title={`Filter the page to ${row.title}`}
                        className={`min-w-0 truncate text-left transition hover:text-ctp-blue ${
                          isSelected ? 'font-semibold text-ctp-blue' : 'text-ctp-text'
                        }`}
                        onClick={() => onSelectTitle(row.title)}
                      >
                        {row.title}
                      </button>
                      <button
                        type="button"
                        aria-label={`Open ${row.title}`}
                        title="Open this title"
                        className="shrink-0 text-ctp-overlay2 transition hover:text-ctp-blue"
                        onClick={() => onOpenTitle(row)}
                      >
                        ↗
                      </button>
                    </div>
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    <div>{formatWatchTime(row.watchTimeMin)}</div>
                    <div className="ml-auto mt-1 h-1 w-20 overflow-hidden rounded-full bg-ctp-surface1/60">
                      <div
                        className="ml-auto h-full rounded-full bg-ctp-blue"
                        style={{
                          width: `${maxWatchTime > 0 ? (row.watchTimeMin / maxWatchTime) * 100 : 0}%`,
                        }}
                      />
                    </div>
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.videos)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-text tabular-nums">
                    {formatNumber(row.episodesFinished)}
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
                    {row.lookupsPerHundred === null ? '—' : row.lookupsPerHundred.toFixed(1)}
                  </td>
                  <td className="px-2 py-2 text-right text-ctp-subtext0 tabular-nums">
                    {formatDateRange(row.firstWatched, row.lastWatched)}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
