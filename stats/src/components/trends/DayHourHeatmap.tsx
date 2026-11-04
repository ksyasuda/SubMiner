import { formatHours } from '../../lib/formatters';

interface DayHourHeatmapProps {
  /** Active minutes as `[dayOfWeek][hour]`, with 0 = Sunday. */
  grid: number[][];
}

const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const DAY_ORDER = [1, 2, 3, 4, 5, 6, 0];
// The day starts at 06:00 so an evening that runs past midnight stays one block
// instead of being split across both ends of the chart.
const FIRST_HOUR = 6;
const HOUR_LABEL_EVERY = 3;
const MIN_CELL_OPACITY = 0.18;

export const HEATMAP_HOURS = Array.from({ length: 24 }, (_, index) => (index + FIRST_HOUR) % 24);

function formatHourLabel(hour: number): string {
  return `${String(hour).padStart(2, '0')}:00`;
}

/**
 * Square-root scale, so a few very long sessions don't wash out the rest of the
 * grid. Returns 0 for an empty cell, which is drawn as background rather than color.
 */
export function heatmapCellOpacity(minutes: number, maxMinutes: number): number {
  if (minutes <= 0 || maxMinutes <= 0) return 0;
  return MIN_CELL_OPACITY + (1 - MIN_CELL_OPACITY) * Math.sqrt(minutes / maxMinutes);
}

export function DayHourHeatmap({ grid }: DayHourHeatmapProps) {
  const maxMinutes = Math.max(0, ...grid.flat());
  const columns = `2.25rem repeat(24, minmax(0, 1fr)) 3rem`;

  return (
    <div className="col-span-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4">
      <h3 className="mb-2 text-xs font-semibold text-ctp-text">Watch Time by Day and Hour</h3>
      {maxMinutes === 0 ? (
        <div className="text-xs text-ctp-overlay2">No data</div>
      ) : (
        <div className="grid items-center gap-0.5" style={{ gridTemplateColumns: columns }}>
          {DAY_ORDER.map((day) => {
            const hours = grid[day] ?? [];
            const total = hours.reduce((sum, minutes) => sum + minutes, 0);
            return (
              <div key={day} className="contents">
                <div className="text-[10px] text-ctp-subtext0">{DAY_NAMES[day]}</div>
                {HEATMAP_HOURS.map((hour) => {
                  const minutes = hours[hour] ?? 0;
                  const opacity = heatmapCellOpacity(minutes, maxMinutes);
                  return (
                    <div
                      key={hour}
                      title={`${DAY_NAMES[day]} ${formatHourLabel(hour)}: ${formatHours(minutes)}`}
                      className={`h-6 rounded-sm ${opacity === 0 ? 'bg-ctp-surface1/30' : 'bg-ctp-blue'}`}
                      style={opacity === 0 ? undefined : { opacity }}
                    />
                  );
                })}
                <div className="text-right text-[10px] text-ctp-subtext0 tabular-nums">
                  {formatHours(total)}
                </div>
              </div>
            );
          })}
          <div />
          {HEATMAP_HOURS.map((hour, index) => (
            <div key={hour} className="text-[10px] whitespace-nowrap text-ctp-overlay2">
              {index % HOUR_LABEL_EVERY === 0 ? formatHourLabel(hour) : ''}
            </div>
          ))}
          <div />
        </div>
      )}
    </div>
  );
}
