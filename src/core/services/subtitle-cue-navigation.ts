import type { SubtitleCue } from './subtitle-cue-parser';

const CUE_START_GROUP_TOLERANCE_SECONDS = 0.05;
const CUE_BOUNDARY_SEEK_OFFSET_SECONDS = 0.08;
const CUE_END_GUARD_SECONDS = 0.01;

type CueGroup = {
  startTime: number;
  endTime: number;
  cue: SubtitleCue;
};

function isValidCue(cue: SubtitleCue): boolean {
  return (
    Number.isFinite(cue.startTime) && Number.isFinite(cue.endTime) && cue.endTime > cue.startTime
  );
}

function groupCueBoundaries(cues: readonly SubtitleCue[]): CueGroup[] {
  const sorted = cues.filter(isValidCue).sort((left, right) => {
    return left.startTime - right.startTime || left.endTime - right.endTime;
  });
  const groups: CueGroup[] = [];

  for (const cue of sorted) {
    const current = groups.at(-1);
    if (current && cue.startTime - current.startTime <= CUE_START_GROUP_TOLERANCE_SECONDS) {
      current.endTime = Math.max(current.endTime, cue.endTime);
      continue;
    }
    groups.push({ startTime: cue.startTime, endTime: cue.endTime, cue });
  }

  return groups;
}

/** A small offset avoids asking mpv to render exactly on a subtitle boundary. */
export function subtitleCueSeekTime(cue: SubtitleCue): number {
  return Math.max(
    cue.startTime,
    Math.min(cue.endTime - CUE_END_GUARD_SECONDS, cue.startTime + CUE_BOUNDARY_SEEK_OFFSET_SECONDS),
  );
}

/**
 * Choose a stable point inside a selected cue. Karaoke lines can overlap while the
 * previous line animates out, so a sidebar selection should clear that overlap when
 * the selected cue has enough time remaining.
 */
export function subtitleCueListSeekTime(
  cues: readonly SubtitleCue[],
  selectedCue: SubtitleCue,
): number {
  const groups = groupCueBoundaries(cues);
  const selectedGroupIndex = groups.findIndex(
    (group) =>
      selectedCue.startTime >= group.startTime &&
      selectedCue.startTime - group.startTime <= CUE_START_GROUP_TOLERANCE_SECONDS,
  );
  const previousGroupEndTime =
    selectedGroupIndex > 0 ? groups[selectedGroupIndex - 1]?.endTime : undefined;
  if (previousGroupEndTime === undefined || previousGroupEndTime <= selectedCue.startTime) {
    return subtitleCueSeekTime(selectedCue);
  }

  return Math.max(
    selectedCue.startTime,
    Math.min(
      selectedCue.endTime - CUE_END_GUARD_SECONDS,
      previousGroupEndTime + CUE_BOUNDARY_SEEK_OFFSET_SECONDS,
    ),
  );
}

/**
 * Translate mpv subtitle-line navigation onto parsed cues. Generated ASS karaoke can
 * contain hundreds of subtitle events for one visible line, while the parsed list has
 * already collapsed those events into the authored lines the user expects to navigate.
 */
export function resolveSanitizedSubtitleSeekCommand(
  command: readonly (string | number)[],
  cues: readonly SubtitleCue[],
  currentTimeSec: number,
): (string | number)[] | null {
  if (
    command.length < 2 ||
    command[0] !== 'sub-seek' ||
    (command[1] !== -1 && command[1] !== 1) ||
    !Number.isFinite(currentTimeSec)
  ) {
    return null;
  }

  const groups = groupCueBoundaries(cues);
  if (groups.length === 0) {
    return null;
  }

  let activeIndex = -1;
  for (const [index, group] of groups.entries()) {
    if (group.startTime <= currentTimeSec && group.endTime > currentTimeSec) {
      activeIndex = index;
    }
  }

  let destination: CueGroup | undefined;
  if (command[1] === 1) {
    destination =
      activeIndex >= 0
        ? groups[activeIndex + 1]
        : groups.find((group) => group.startTime > currentTimeSec);
  } else if (activeIndex >= 0) {
    destination = groups[activeIndex - 1];
  } else {
    for (let index = groups.length - 1; index >= 0; index -= 1) {
      const group = groups[index]!;
      if (group.startTime < currentTimeSec) {
        destination = group;
        break;
      }
    }
  }

  if (!destination) {
    return null;
  }
  return ['seek', subtitleCueSeekTime(destination.cue), 'absolute+exact'];
}
