/*
 * Duplicate/animation-burst collapsing for parsed subtitle cues.
 *
 * Split out of the cue parser so the parsing rules and the "is this run one animation?"
 * heuristics can be read -- and tested -- on their own. The parser owns the cue shape;
 * this module only decides which cues survive.
 */

import { hasAssTemporalOverride, isAnimatedAssEffectKind } from './ass-text';
import {
  ANIMATION_FRAME_MAX_SECONDS,
  DUPLICATE_CUE_GAP_TOLERANCE_SECONDS,
  MIN_BURST_EVENTS,
  MIN_TAGGED_BURST_FRAMES,
  MIN_TIMING_ONLY_FRAMES,
  TIMING_ONLY_FRAME_MAX_SECONDS,
} from './subtitle-burst-constants';
import type {
  AnnotatedSubtitleCue,
  SubtitleCue,
  SubtitleSourceFormat,
} from './subtitle-cue-parser';

function cueKey(cue: SubtitleCue): string {
  return `${cue.startTime}|${cue.endTime}|${cue.text}`;
}

/**
 * Identical text over an identical span is redundant however it was authored -- most
 * often a layered ASS event stacking a shadow copy under the visible one.
 */
function collapseExactDuplicates(cues: AnnotatedSubtitleCue[]): AnnotatedSubtitleCue[] {
  const seen = new Set<string>();
  return cues.filter((cue) => {
    const key = cueKey(cue);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function countFramesShorterThan(run: AnnotatedSubtitleCue[], maxSeconds: number): number {
  return run.filter((cue) => cue.endTime - cue.startTime < maxSeconds).length;
}

/**
 * Evidence that a run of ASS events is one animation rather than several authored lines.
 * A static tag says nothing on its own -- three events sharing one `\clip(...)` are three
 * signs -- so the tag has to be temporal by nature (`\t`, `\move`, karaoke timing, or
 * anything wrapped in `\t(...)`), an animated `Effect` column, or a value that actually
 * changes from event to event, which is how per-frame typesetting is authored.
 */
export function hasAssAnimationEvidence(run: AnnotatedSubtitleCue[]): boolean {
  if (run.every((cue) => hasAssTemporalOverride(cue.overrides))) {
    return true;
  }
  if (run.every((cue) => isAnimatedAssEffectKind(cue.effectKind))) {
    return true;
  }

  const [first] = run;
  const everyEventTypeset = run.every((cue) => cue.overrides.length > 0);
  const signatureChanges = run.some((cue) => cue.overrideSignature !== first!.overrideSignature);
  return everyEventTypeset && signatureChanges;
}

export function isAnimationBurst(
  run: AnnotatedSubtitleCue[],
  format: SubtitleSourceFormat,
): boolean {
  if (run.length < MIN_BURST_EVENTS) {
    return false;
  }

  if (format === 'srt') {
    return (
      run.length >= MIN_TIMING_ONLY_FRAMES &&
      countFramesShorterThan(run, TIMING_ONLY_FRAME_MAX_SECONDS) === run.length
    );
  }

  if (countFramesShorterThan(run, ANIMATION_FRAME_MAX_SECONDS) < MIN_TAGGED_BURST_FRAMES) {
    return false;
  }

  // One animation belongs to one styled, one named source line. Two characters trading
  // the same short word are two styles or two actors, and never merge.
  const [first] = run;
  if (run.some((cue) => cue.style !== first!.style || cue.name !== first!.name)) {
    return false;
  }

  return hasAssAnimationEvidence(run);
}

/**
 * Karaoke and sign typesetting emits one Dialogue event per animation frame, all carrying
 * the same visible text over a contiguous span. Collapse each such run into a single cue.
 *
 * Only runs that look like animation collapse. Two ordinary lines that happen to repeat
 * -- several characters each saying `おはよう` in turn, a positioned sign redrawn with a
 * different fade -- stay separate, because merging them would destroy real mineable lines.
 */
function collapseAnimationBursts(
  cues: AnnotatedSubtitleCue[],
  format: SubtitleSourceFormat,
): AnnotatedSubtitleCue[] {
  const indicesByText = new Map<string, number[]>();
  cues.forEach((cue, index) => {
    const bucket = indicesByText.get(cue.text);
    if (bucket) {
      bucket.push(index);
    } else {
      indicesByText.set(cue.text, [index]);
    }
  });

  const dropped = new Set<number>();
  const extendedEnd = new Map<number, number>();

  for (const indices of indicesByText.values()) {
    if (indices.length < MIN_BURST_EVENTS) {
      continue;
    }

    let runStart = 0;
    while (runStart < indices.length) {
      let runEnd = runStart;
      let chainEnd = cues[indices[runStart]!]!.endTime;

      while (runEnd + 1 < indices.length) {
        const next = cues[indices[runEnd + 1]!]!;
        if (next.startTime > chainEnd + DUPLICATE_CUE_GAP_TOLERANCE_SECONDS) {
          break;
        }
        chainEnd = Math.max(chainEnd, next.endTime);
        runEnd += 1;
      }

      const run = indices.slice(runStart, runEnd + 1).map((index) => cues[index]!);
      if (isAnimationBurst(run, format)) {
        for (let i = runStart + 1; i <= runEnd; i += 1) {
          dropped.add(indices[i]!);
        }
        extendedEnd.set(indices[runStart]!, chainEnd);
      }

      runStart = runEnd + 1;
    }
  }

  if (dropped.size === 0) {
    return cues;
  }

  const merged: AnnotatedSubtitleCue[] = [];
  cues.forEach((cue, index) => {
    if (dropped.has(index)) {
      return;
    }
    const end = extendedEnd.get(index);
    merged.push(end !== undefined && end > cue.endTime ? { ...cue, endTime: end } : cue);
  });
  return merged;
}

/**
 * Collapse redundant cues. Input must already be sorted by non-decreasing `startTime`,
 * ties broken by `endTime` then source `order` -- burst detection chains events by
 * comparing each one against the running end of the events before it, so an unsorted
 * list breaks runs apart and leaves the frames behind.
 */
export function mergeDuplicateCues(
  cues: AnnotatedSubtitleCue[],
  format: SubtitleSourceFormat,
): AnnotatedSubtitleCue[] {
  return collapseAnimationBursts(collapseExactDuplicates(cues), format);
}
