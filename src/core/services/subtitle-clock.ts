/**
 * mpv's `sub-start`/`sub-end` and parsed cues carry the subtitle file's own timestamps
 * (the subtitle clock), but mpv plays each line `sub-delay` seconds later (z/x retiming,
 * `sub-step`). Playback position, seeks, and media cuts are on the video clock, so any
 * value crossing between the two converts here.
 */

function finiteDelay(subDelay: number): number {
  return Number.isFinite(subDelay) ? subDelay : 0;
}

/** Playback position (`time-pos`) on the subtitle clock, comparable with cue times. */
export function toSubtitleClock(videoTimeSec: number, subDelay: number): number {
  return videoTimeSec - finiteDelay(subDelay);
}

/** Subtitle-clock seconds on the video clock, for seeks. */
export function toVideoClock(subtitleTimeSec: number, subDelay: number): number {
  return subtitleTimeSec + finiteDelay(subDelay);
}

/** A subtitle-clock range on the video clock, for mining and media generation. */
export function shiftSubtitleRangeToVideoClock<T extends { startTime: number; endTime: number }>(
  range: T,
  subDelay: number,
): T {
  const delay = finiteDelay(subDelay);
  if (delay === 0) {
    return range;
  }
  return { ...range, startTime: range.startTime + delay, endTime: range.endTime + delay };
}
