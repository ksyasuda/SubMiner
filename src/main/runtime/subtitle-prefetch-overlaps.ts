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
  const singleTexts = new Set(cues.map((cue) => cue.text));
  const seen = new Set<string>();
  const overlapCues: SubtitleCue[] = [];

  for (let i = 0; i + 1 < boundaries.length; i += 1) {
    const startTime = boundaries[i]!;
    const endTime = boundaries[i + 1]!;
    const midpoint = (startTime + endTime) / 2;
    const active = cues.filter((cue) => cue.startTime <= midpoint && cue.endTime > midpoint);
    if (active.length < 2) continue;

    const text = resolvePrimarySubtitleText({
      liveText: active.map((cue) => cue.text).join('\n'),
      currentTimeSec: midpoint,
      cues,
    });
    if (!text.trim() || singleTexts.has(text) || seen.has(text)) continue;
    seen.add(text);
    overlapCues.push({ startTime, endTime, text });
  }

  return overlapCues;
}
