import type { SubtitleCue } from '../../types';

export function selectAutoplayStartupCue(
  cues: SubtitleCue[],
  currentTimeSeconds: number,
  lookaheadSeconds: number,
): SubtitleCue | null {
  const currentTime = Number.isFinite(currentTimeSeconds) ? currentTimeSeconds : 0;
  const lookahead = Math.max(0, Number.isFinite(lookaheadSeconds) ? lookaheadSeconds : 0);
  const latestStartTime = currentTime + lookahead;

  for (const cue of cues) {
    if (!cue.text.trim()) {
      continue;
    }
    if (cue.startTime <= currentTime && cue.endTime > currentTime) {
      return cue;
    }
  }

  for (const cue of cues) {
    if (!cue.text.trim()) {
      continue;
    }
    if (cue.startTime >= currentTime && cue.startTime <= latestStartTime) {
      return cue;
    }
  }

  return null;
}
