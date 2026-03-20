import { useState } from 'react';
import type { StreakCalendarPoint } from '../../lib/dashboard-data';

interface StreakCalendarProps {
  data: StreakCalendarPoint[];
}

function intensityClass(value: number): string {
  if (value === 0) return 'bg-ctp-surface0';
  if (value <= 30) return 'bg-ctp-green/30';
  if (value <= 60) return 'bg-ctp-green/60';
  return 'bg-ctp-green';
}

const DAY_LABELS = ['Mon', '', 'Wed', '', 'Fri', '', ''];

export function StreakCalendar({ data }: StreakCalendarProps) {
  const [tooltip, setTooltip] = useState<{ x: number; y: number; text: string } | null>(null);

  const lookup = new Map(data.map((d) => [d.date, d.value]));

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const endDate = new Date(today);
  const startDate = new Date(today);
  startDate.setDate(startDate.getDate() - 89);

  const startDow = (startDate.getDay() + 6) % 7;

  const cells: Array<{ date: string; value: number; row: number; col: number }> = [];
  let col = 0;
  let row = startDow;

  const cursor = new Date(startDate);
  while (cursor <= endDate) {
    const dateStr = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, '0')}-${String(cursor.getDate()).padStart(2, '0')}`;
    cells.push({ date: dateStr, value: lookup.get(dateStr) ?? 0, row, col });

    row += 1;
    if (row >= 7) {
      row = 0;
      col += 1;
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  const totalCols = col + (row > 0 ? 1 : 0);

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4">
      <h3 className="text-sm font-semibold text-ctp-text mb-3">Activity (90 days)</h3>
      <div className="relative flex gap-1">
        <div className="flex flex-col gap-1 text-[10px] text-ctp-overlay2 pr-1 shrink-0">
          {DAY_LABELS.map((label, i) => (
            <div key={i} className="h-3 flex items-center leading-none">
              {label}
            </div>
          ))}
        </div>
        <div
          className="grid gap-[3px]"
          style={{
            gridTemplateColumns: `repeat(${totalCols}, 12px)`,
            gridTemplateRows: 'repeat(7, 12px)',
          }}
        >
          {cells.map((cell) => (
            <div
              key={cell.date}
              className={`w-3 h-3 rounded-sm ${intensityClass(cell.value)} cursor-default`}
              style={{ gridRow: cell.row + 1, gridColumn: cell.col + 1 }}
              onMouseEnter={(e) => {
                const rect = e.currentTarget.getBoundingClientRect();
                setTooltip({
                  x: rect.left + rect.width / 2,
                  y: rect.top - 4,
                  text: `${cell.date}: ${Math.round(cell.value * 100) / 100}m`,
                });
              }}
              onMouseLeave={() => setTooltip(null)}
            />
          ))}
        </div>
        {tooltip && (
          <div
            className="fixed z-50 px-2 py-1 text-xs bg-ctp-crust text-ctp-text rounded shadow-lg pointer-events-none -translate-x-1/2 -translate-y-full"
            style={{ left: tooltip.x, top: tooltip.y }}
          >
            {tooltip.text}
          </div>
        )}
      </div>
    </div>
  );
}
