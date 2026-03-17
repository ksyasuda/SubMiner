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
  return new Date(epochDay * 86_400_000);
}

export function localDayFromMs(ms: number): number {
  const d = new Date(ms);
  const localMidnight = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  return Math.floor(localMidnight / 86_400_000);
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
  const includeYear = date.getFullYear() !== new Date().getFullYear();
  return date.toLocaleDateString(undefined, {
    month: 'long',
    day: 'numeric',
    ...(includeYear ? { year: 'numeric' } : {}),
  });
}
