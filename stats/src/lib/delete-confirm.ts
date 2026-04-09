export function confirmSessionDelete(): boolean {
  return globalThis.confirm('Delete this session and all associated data?');
}

export function confirmDayGroupDelete(dayLabel: string, count: number): boolean {
  return globalThis.confirm(
    `Delete all ${count} session${count === 1 ? '' : 's'} from ${dayLabel} and all associated data?`,
  );
}

export function confirmAnimeGroupDelete(title: string, count: number): boolean {
  return globalThis.confirm(
    `Delete all ${count} session${count === 1 ? '' : 's'} for "${title}" and all associated data?`,
  );
}

export function confirmEpisodeDelete(title: string): boolean {
  return globalThis.confirm(`Delete "${title}" and all its sessions?`);
}

export function confirmBucketDelete(title: string, count: number): boolean {
  return globalThis.confirm(
    `Delete all ${count} session${count === 1 ? '' : 's'} of "${title}" from this day?`,
  );
}
