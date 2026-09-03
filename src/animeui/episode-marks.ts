/**
 * Which episodes a manual watch mark applies to.
 *
 * Sources list newest first, so "everything below" the row you right-clicked is
 * the back catalogue: the natural shape of "I have watched up to here". The
 * span is taken from the full list rather than from what the filter leaves,
 * because a filter narrows what you are looking at, not what you have watched.
 */

export type MarkScope = 'one' | 'below';

export function episodesInScope<T>(episodes: T[], index: number, scope: MarkScope): T[] {
  if (index < 0 || index >= episodes.length) return [];
  return scope === 'one' ? [episodes[index]!] : episodes.slice(index);
}

/** "3 episodes" / "1 episode", for the status line after a bulk mark. */
export function describeMarkCount(count: number, watched: boolean): string {
  const noun = count === 1 ? 'episode' : 'episodes';
  return watched ? `Marked ${count} ${noun} watched` : `Cleared the watch mark on ${count} ${noun}`;
}
