import type { AnimeBrowserQueueEntry } from '../types/anime-browser';

/**
 * Where each queued episode sits in line, keyed by source and episode url.
 *
 * The position is counted across the whole queue rather than within one anime:
 * "3rd up" has to mean the same thing on every detail page, or the number is
 * worse than no number at all.
 */
export function queuePositions(entries: AnimeBrowserQueueEntry[]): Map<string, number> {
  const positions = new Map<string, number>();
  entries.forEach((entry, index) => {
    const key = queueKey(entry.sourceId, entry.episodeUrl);
    // A duplicate should never reach the renderer, but if one did, the earlier
    // position is the one that will actually play.
    if (!positions.has(key)) positions.set(key, index + 1);
  });
  return positions;
}

export function queueKey(sourceId: string, episodeUrl: string): string {
  return `${sourceId} ${episodeUrl}`;
}

/** "Next up" reads better than "#1" for the episode about to play. */
export function describeQueuePosition(position: number): string {
  return position === 1 ? 'next up' : `#${position} in queue`;
}
