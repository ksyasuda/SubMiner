import { formatDuration, formatRelativeDate, formatNumber } from '../../lib/formatters';
import type { SessionSummary } from '../../types/stats';

interface MediaSessionListProps {
  sessions: SessionSummary[];
}

export function MediaSessionList({ sessions }: MediaSessionListProps) {
  if (sessions.length === 0) {
    return <div className="text-sm text-ctp-overlay2">No sessions recorded</div>;
  }

  return (
    <div className="space-y-2">
      <h3 className="text-sm font-semibold text-ctp-text">Session History</h3>
      {sessions.map((s) => (
        <div
          key={s.sessionId}
          className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-3 flex items-center justify-between"
        >
          <div className="min-w-0">
            <div className="text-sm text-ctp-text">
              {formatRelativeDate(s.startedAtMs)} · {formatDuration(s.activeWatchedMs)} active
            </div>
          </div>
          <div className="flex gap-4 text-xs text-center shrink-0">
            <div>
              <div className="text-ctp-green font-medium">{formatNumber(s.cardsMined)}</div>
              <div className="text-ctp-overlay2">cards</div>
            </div>
            <div>
              <div className="text-ctp-mauve font-medium">{formatNumber(s.wordsSeen)}</div>
              <div className="text-ctp-overlay2">words</div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
