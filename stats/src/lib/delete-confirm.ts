export function confirmSessionDelete(): boolean {
  return globalThis.confirm('Delete this session and all associated data?');
}

export function confirmEpisodeDelete(title: string): boolean {
  return globalThis.confirm(`Delete "${title}" and all its sessions?`);
}
