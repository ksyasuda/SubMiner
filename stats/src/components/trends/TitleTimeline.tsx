import { useMemo, useState } from 'react';
import type { LibrarySummaryRow, TrendPerAnimePoint } from '../../types/stats';
import { epochDayToDate, formatHours, formatNumber } from '../../lib/formatters';
import { heatmapCellOpacity } from './DayHourHeatmap';
import { buildTitleTimeline } from './title-timeline';

interface TitleTimelineProps {
  /** Titles to show, most watched first. */
  rows: LibrarySummaryRow[];
  points: TrendPerAnimePoint[];
  window: { startDay: number; endDay: number };
  selectedTitle: string | null;
  onSelectTitle: (title: string) => void;
  onOpenTitle: (row: LibrarySummaryRow) => void;
}

const DEFAULT_VISIBLE_ROWS = 12;

function formatCellDate(epochDay: number): string {
  return epochDayToDate(epochDay).toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  });
}

/** One row per title, shaded on the days it was watched, so viewing bursts are visible at a glance. */
export function TitleTimeline({
  rows,
  points,
  window,
  selectedTitle,
  onSelectTitle,
  onOpenTitle,
}: TitleTimelineProps) {
  const [showAll, setShowAll] = useState(false);
  const model = useMemo(() => buildTitleTimeline(points, window), [points, window]);
  const visibleRows = showAll ? rows : rows.slice(0, DEFAULT_VISIBLE_ROWS);

  return (
    <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
      <h3 className="mb-2 text-xs font-semibold text-ctp-text">
        What You Watched and When
        {model.columnDays > 1 ? (
          <span className="ml-2 font-normal text-ctp-overlay2">One column per week</span>
        ) : null}
      </h3>
      {rows.length === 0 ? (
        <div className="text-xs text-ctp-overlay2">No titles in this range.</div>
      ) : (
        <>
          <div className="grid grid-cols-[minmax(7rem,14rem)_minmax(0,1fr)_2.5rem_2.5rem] items-center gap-x-2 gap-y-px">
            <div />
            <div className="relative h-4">
              {model.ticks.map((tick) => (
                <span
                  key={tick.leftPercent}
                  className="absolute top-0 text-[10px] whitespace-nowrap text-ctp-overlay2"
                  style={{ left: `${tick.leftPercent}%` }}
                >
                  {tick.label}
                </span>
              ))}
            </div>
            <div className="text-right text-[10px] text-ctp-overlay2">Time</div>
            <div className="text-right text-[10px] text-ctp-overlay2">Cards</div>
            {visibleRows.map((row) => {
              const isSelected = row.title === selectedTitle;
              return (
                <div key={row.title} className="contents">
                  <div className="flex min-w-0 items-center gap-1">
                    <button
                      type="button"
                      aria-pressed={isSelected}
                      title={`Filter the page to ${row.title}`}
                      className={`min-w-0 truncate text-left text-[11px] transition hover:text-ctp-blue ${
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
                      className="shrink-0 text-[11px] text-ctp-overlay2 transition hover:text-ctp-blue"
                      onClick={() => onOpenTitle(row)}
                    >
                      ↗
                    </button>
                  </div>
                  <div
                    className={`relative h-4 rounded-sm ${isSelected ? 'bg-ctp-surface2/60' : 'bg-ctp-surface1/30'}`}
                  >
                    {model.ticks.map((tick) => (
                      <span
                        key={tick.leftPercent}
                        className="absolute inset-y-0 w-px bg-ctp-surface2/60"
                        style={{ left: `${tick.leftPercent}%` }}
                      />
                    ))}
                    {(model.cellsByTitle.get(row.title) ?? []).map((cell) => (
                      <span
                        key={cell.startDay}
                        title={`${row.title}\n${model.columnDays > 1 ? 'Week of ' : ''}${formatCellDate(cell.startDay)}: ${formatHours(cell.minutes)}`}
                        className="absolute inset-y-px min-w-[2px] rounded-[1px] bg-ctp-blue"
                        style={{
                          left: `${cell.leftPercent}%`,
                          width: `${cell.widthPercent}%`,
                          opacity: heatmapCellOpacity(cell.minutes, model.maxCellMinutes),
                        }}
                      />
                    ))}
                  </div>
                  <div className="text-right text-[10px] text-ctp-subtext0 tabular-nums">
                    {formatHours(row.watchTimeMin)}
                  </div>
                  <div className="text-right text-[10px] text-ctp-subtext0 tabular-nums">
                    {formatNumber(row.cards)}
                  </div>
                </div>
              );
            })}
          </div>
          {rows.length > DEFAULT_VISIBLE_ROWS ? (
            <button
              type="button"
              className="mt-2 text-[11px] text-ctp-overlay2 transition hover:text-ctp-blue"
              onClick={() => setShowAll((value) => !value)}
            >
              {showAll ? 'Show fewer' : `Show all ${formatNumber(rows.length)} titles`}
            </button>
          ) : null}
        </>
      )}
    </div>
  );
}
