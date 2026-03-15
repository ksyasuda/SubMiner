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
  if (diffMs < 60_000) return 'just now';
  const diffMin = Math.floor(diffMs / 60_000);
  if (diffMin < 60) return `${diffMin}m ago`;
  const diffHours = Math.floor(diffMs / 3_600_000);
  if (diffHours < 24) return `${diffHours}h ago`;
  const diffDays = Math.floor(diffMs / 86_400_000);
  if (diffDays < 2) return 'Yesterday';
  if (diffDays < 7) return `${diffDays}d ago`;
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
