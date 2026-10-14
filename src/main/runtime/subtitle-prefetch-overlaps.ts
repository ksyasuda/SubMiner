import type { SubtitleCue } from '../../types';
import { resolvePrimarySubtitleText } from './primary-subtitle-text';

/**
 * While cues overlap, mpv publishes their texts joined as one `sub-text`, and the live
 * path tokenizes and caches that combined line as its own entry. Prefetching single cues
 * never warms it, so overlapping lines (usually the long ones) always miss the cache.
 *
 * Returns one synthetic cue per overlap span carrying the text the live path resolves
 * for it, so the prefetcher can warm those lines in timeline order with the rest.
 */
export function buildOverlapPrefetchCues(cues: readonly SubtitleCue[]): SubtitleCue[] {
  const boundaries = [...new Set(cues.flatMap((cue) => [cue.startTime, cue.endTime]))].sort(
    (a, b) => a - b,
  );
  const indices = cues.map((_, index) => index);
  const byStart = [...indices].sort((a, b) => cues[a]!.startTime - cues[b]!.startTime);
  const byEnd = [...indices].sort((a, b) => cues[a]!.endTime - cues[b]!.endTime);
  const singleTexts = new Set(cues.map((cue) => cue.text));
  const seen = new Set<string>();
  const overlapCues: SubtitleCue[] = [];

  // Sweep the boundaries once, keeping the cues active over [startTime, endTime). Every
  // cue edge is a boundary, so a cue is active for a whole interval or not at all.
  const active = new Set<number>();
  let nextStart = 0;
  let nextEnd = 0;
  for (let i = 0; i + 1 < boundaries.length; i += 1) {
    const startTime = boundaries[i]!;
    const endTime = boundaries[i + 1]!;
    for (; nextEnd < byEnd.length && cues[byEnd[nextEnd]!]!.endTime <= startTime; nextEnd += 1) {
      active.delete(byEnd[nextEnd]!);
    }
    for (
      ;
      nextStart < byStart.length && cues[byStart[nextStart]!]!.startTime <= startTime;
      nextStart += 1
    ) {
      const index = byStart[nextStart]!;
      // Empty or inverted cues are never on screen.
      if (cues[index]!.endTime > startTime) active.add(index);
    }
    if (active.size < 2) continue;

    // mpv joins simultaneous cues in cue-list order.
    const liveText = [...active]
      .sort((a, b) => a - b)
      .map((index) => cues[index]!.text)
      .join('\n');
    const text = resolvePrimarySubtitleText({
      liveText,
      currentTimeSec: (startTime + endTime) / 2,
      cues,
    });
    if (!text.trim() || singleTexts.has(text) || seen.has(text)) continue;
    seen.add(text);
    overlapCues.push({ startTime, endTime, text });
  }

  return overlapCues;
}
