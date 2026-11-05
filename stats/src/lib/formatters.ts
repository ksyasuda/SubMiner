export function formatDuration(ms: number): string {
  const totalMin = Math.round(ms / 60_000);
  if (totalMin < 60) return `${totalMin}m`;
  const hours = Math.floor(totalMin / 60);
  const mins = totalMin % 60;
  return mins > 0 ? `${hours}h ${mins}m` : `${hours}h`;
}

export function formatNumber(n: number): string {
  return n.toLocaleString();
}

const COMPACT_NUMBER = new Intl.NumberFormat(undefined, {
  notation: 'compact',
  maximumFractionDigits: 1,
});

/** Short form for tight spaces like chart axes: 24000 becomes 24K. */
export function formatCompactNumber(n: number): string {
  return COMPACT_NUMBER.format(n);
}

/** Watch time for headline numbers: minutes below an hour, then hours to one decimal. */
export function formatHours(minutes: number): string {
  return minutes < 60 ? `${Math.round(minutes)}m` : `${(minutes / 60).toFixed(1)}h`;
}

export function formatPercent(ratio: number | null): string {
  if (ratio == null) return '\u2014';
  return `${Math.round(ratio * 100)}%`;
}

export function formatRelativeDate(ms: number): string {
  const now = Date.now();
  const diffMs = now - ms;
  if (diffMs <= 0) return 'just now';

  const nowDay = localDayFromMs(now);
  const sessionDay = localDayFromMs(ms);
  const dayDiff = nowDay - sessionDay;

  if (dayDiff <= 0) {
    if (diffMs < 60_000) return 'just now';
    const diffMin = Math.floor(diffMs / 60_000);
    if (diffMin < 60) return `${diffMin}m ago`;
    const diffHours = Math.floor(diffMs / 3_600_000);
    return `${diffHours}h ago`;
  }

  if (dayDiff === 1) return 'Yesterday';
  if (dayDiff < 7) return `${dayDiff}d ago`;
  return new Date(ms).toLocaleDateString();
}

export function epochDayToDate(epochDay: number): Date {
  const utcDate = new Date(epochDay * 86_400_000);
  return new Date(utcDate.getUTCFullYear(), utcDate.getUTCMonth(), utcDate.getUTCDate());
}

/** `YYYY-MM-DD` for a date input. */
export function epochDayToIsoDate(epochDay: number): string {
  return new Date(epochDay * 86_400_000).toISOString().slice(0, 10);
}

/** Epoch day for a date input value, or null when the input is empty, invalid, or before 1970. */
export function isoDateToEpochDay(value: string): number | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return null;
  const year = Number(match[1]);
  if (year < 1970) return null;
  const ms = Date.UTC(year, Number(match[2]) - 1, Number(match[3]));
  return Number.isNaN(ms) ? null : ms / 86_400_000;
}

/** The local calendar day as an epoch day, the same numbering the stats server uses. */
export function localDayFromMs(ms: number): number {
  const d = new Date(ms);
  return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86_400_000;
}

export function todayLocalDay(): number {
  return localDayFromMs(Date.now());
}

// Immersion tracker stores word/kanji first_seen/last_seen as epoch seconds.
// Older fixtures or callers may still pass ms, so normalize defensively.
export function epochMsFromDbTimestamp(ts: number): number {
  if (!Number.isFinite(ts)) return 0;
  return ts < 10_000_000_000 ? Math.round(ts * 1000) : Math.round(ts);
}

export function formatSessionDayLabel(sessionStartedAtMs: number): string {
  const today = todayLocalDay();
  const day = localDayFromMs(sessionStartedAtMs);

  if (day === today) return 'Today';
  if (day === today - 1) return 'Yesterday';

  const date = new Date(sessionStartedAtMs);
  const includeYear = date.getFullYear() !== new Date(Date.now()).getFullYear();
  return date.toLocaleDateString(undefined, {
    month: 'long',
    day: 'numeric',
    ...(includeYear ? { year: 'numeric' } : {}),
  });
}

/** Second-precise duration for short spans ("42s", "6m 12s", "1h 3m"). */
export function formatClockDuration(ms: number): string {
  const totalSec = Math.max(0, Math.round(ms / 1000));
  if (totalSec < 60) return `${totalSec}s`;
  const totalMin = Math.floor(totalSec / 60);
  if (totalMin < 60) {
    const sec = totalSec % 60;
    return sec > 0 ? `${totalMin}m ${sec}s` : `${totalMin}m`;
  }
  const hours = Math.floor(totalMin / 60);
  const mins = totalMin % 60;
  return mins > 0 ? `${hours}h ${mins}m` : `${hours}h`;
}
