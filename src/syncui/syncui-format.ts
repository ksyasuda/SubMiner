import type { SyncMergeSummary } from '../shared/sync/sync-events';

export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${Math.round(bytes)} B`;
  const units = ['KB', 'MB', 'GB', 'TB'];
  let value = bytes / 1024;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(1)} ${units[unitIndex]}`;
}

export function formatRelativeTime(atMs: number | null, nowMs: number): string {
  if (atMs === null) return 'never';
  const elapsed = Math.max(0, nowMs - atMs);
  if (elapsed < 60_000) return 'just now';
  if (elapsed < 3_600_000) return `${Math.floor(elapsed / 60_000)} min ago`;
  if (elapsed < 86_400_000) return `${Math.floor(elapsed / 3_600_000)} h ago`;
  return `${Math.floor(elapsed / 86_400_000)} d ago`;
}

export interface MergeCountLine {
  label: string;
  value: number;
}

const MERGE_COUNT_FIELDS: Array<{ key: keyof SyncMergeSummary; label: string }> = [
  { key: 'sessionsMerged', label: 'Sessions merged' },
  { key: 'sessionsAlreadyPresent', label: 'Already present' },
  { key: 'activeSessionsSkipped', label: 'Active skipped' },
  { key: 'animeAdded', label: 'Series added' },
  { key: 'videosAdded', label: 'Videos added' },
  { key: 'wordsAdded', label: 'Words added' },
  { key: 'kanjiAdded', label: 'Kanji added' },
  { key: 'subtitleLinesAdded', label: 'Subtitle lines' },
  { key: 'excludedWordsAdded', label: 'Excluded words' },
  { key: 'dailyRollupsCopied', label: 'Daily rollups' },
  { key: 'monthlyRollupsCopied', label: 'Monthly rollups' },
  { key: 'rollupGroupsRecomputed', label: 'Rollups recomputed' },
];

// "Sessions merged" always shows (even at 0, it is the headline number);
// everything else appears only when non-zero.
export function summarizeMergeCounts(summary: SyncMergeSummary): MergeCountLine[] {
  return MERGE_COUNT_FIELDS.filter(
    (field) => field.key === 'sessionsMerged' || (summary[field.key] ?? 0) > 0,
  ).map((field) => ({ label: field.label, value: summary[field.key] ?? 0 }));
}
