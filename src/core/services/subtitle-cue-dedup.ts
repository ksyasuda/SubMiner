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
 * often a layered ASS event stacking a shadow copy under the visible one. When one of
 * the duplicates is a recovered canonical cue, that copy survives: dropping it would
 * strip the `source` marker and animation envelope the live overlay substitutes on.
 */
function collapseExactDuplicates(cues: AnnotatedSubtitleCue[]): AnnotatedSubtitleCue[] {
  const survivorByKey = new Map<string, AnnotatedSubtitleCue>();
  const keysInOrder: string[] = [];
  for (const cue of cues) {
    const key = cueKey(cue);
    const existing = survivorByKey.get(key);
    if (!existing) {
      survivorByKey.set(key, cue);
      keysInOrder.push(key);
    } else if (!existing.source && cue.source) {
      survivorByKey.set(key, cue);
    }
  }
  return keysInOrder.map((key) => survivorByKey.get(key)!);
}

const SPATIAL_ASS_OVERRIDE_COMMANDS = new Set([
  'a',
  'an',
  'clip',
  'iclip',
  'move',
  'org',
  'pbo',
  'pos',
  'q',
]);

interface RepeatedPhaseRun {
  cues: AnnotatedSubtitleCue[];
  indices: number[];
}

// A changing override signature alone is weak: two ordinary repeats restyled with
// different colors look identical to a phase pair. Real phase redraws carry a styling
// stack over a full lyric line, and they exist to move a color/highlight boundary
// *within* the line -- so every event also has an override block after visible text
// began. An ordinary restyled repeat carries only a leading block and stays separate.
const MIN_PHASE_EVIDENCE_OVERRIDES = 2;
const MIN_PHASE_TEXT_LENGTH = 4;

function hasMidLineOverrideBlock(rawText: string): boolean {
  let sawVisibleText = false;
  for (let i = 0; i < rawText.length; i += 1) {
    if (rawText[i] === '{') {
      const close = rawText.indexOf('}', i);
      if (close === -1) {
        // Unclosed brace renders as literal text; nothing after it is markup.
        return false;
      }
      if (sawVisibleText) {
        return true;
      }
      i = close;
    } else if (!/\s/.test(rawText[i]!)) {
      sawVisibleText = true;
    }
  }
  return false;
}

function assStyleKey(cue: AnnotatedSubtitleCue): string {
  return `${cue.style}\0${cue.name}\0${cue.layer}`;
}

function spatialOverrideSignature(cue: AnnotatedSubtitleCue): string {
  return cue.overrides
    .filter((command) => SPATIAL_ASS_OVERRIDE_COMMANDS.has(command.name.toLowerCase()))
    .map((command) => `${command.name.toLowerCase()}(${command.args})`)
    .join('|');
}

function hasStableSpatialOverrides(run: readonly AnnotatedSubtitleCue[]): boolean {
  const firstSignature = spatialOverrideSignature(run[0]!);
  return run.every((cue) => spatialOverrideSignature(cue) === firstSignature);
}

function hasDirectPhaseEvidence(run: readonly AnnotatedSubtitleCue[]): boolean {
  // Phases redraw one authored line in place. Whatever the animation evidence, a run
  // whose spatial placement changes is separate authored occurrences -- two flush
  // same-text `\move` signs at different coordinates must never merge.
  if (!hasStableSpatialOverrides(run)) {
    return false;
  }
  if (run.every((cue) => hasAssTemporalOverride(cue.overrides))) {
    return true;
  }
  if (run.every((cue) => isAnimatedAssEffectKind(cue.effectKind))) {
    return true;
  }

  const [first] = run;
  return (
    first!.text.replace(/\s+/gu, '').length >= MIN_PHASE_TEXT_LENGTH &&
    run.every(
      (cue) =>
        cue.overrides.length >= MIN_PHASE_EVIDENCE_OVERRIDES &&
        hasMidLineOverrideBlock(cue.rawText),
    ) &&
    run.some((cue) => cue.overrideSignature !== first!.overrideSignature)
  );
}

function collectRepeatedPhaseRuns(cues: AnnotatedSubtitleCue[]): RepeatedPhaseRun[] {
  const runs: RepeatedPhaseRun[] = [];
  let start = 0;

  while (start < cues.length) {
    const first = cues[start]!;
    const styleKey = assStyleKey(first);
    let end = start;

    while (end + 1 < cues.length) {
      const current = cues[end]!;
      const next = cues[end + 1]!;
      const isFlush =
        Math.abs(next.startTime - current.endTime) <= DUPLICATE_CUE_GAP_TOLERANCE_SECONDS;
      if (
        first.source !== undefined ||
        next.source !== undefined ||
        next.text !== first.text ||
        assStyleKey(next) !== styleKey ||
        !isFlush
      ) {
        break;
      }
      end += 1;
    }

    if (end > start) {
      const indices = Array.from({ length: end - start + 1 }, (_, offset) => start + offset);
      runs.push({
        cues: indices.map((index) => cues[index]!),
        indices,
      });
    }
    start = end + 1;
  }

  return runs;
}

/**
 * Some karaoke scripts redraw one complete lyric for each color/highlight phase. These
 * events last far longer than animation frames, but are still one sidebar/history line.
 * The events must prove themselves through direct animation metadata or changing
 * non-spatial overrides. Plain repeated dialogue and separately positioned signs stay
 * intact.
 */
function collapseAnimatedStylePhases(cues: AnnotatedSubtitleCue[]): AnnotatedSubtitleCue[] {
  const runs = collectRepeatedPhaseRuns(cues);
  if (runs.length === 0) {
    return cues;
  }

  const dropped = new Set<number>();
  const extendedEnd = new Map<number, number>();
  for (const run of runs) {
    if (!hasDirectPhaseEvidence(run.cues)) {
      continue;
    }

    const [firstIndex, ...remainingIndices] = run.indices;
    for (const index of remainingIndices) {
      dropped.add(index);
    }
    extendedEnd.set(firstIndex!, Math.max(...run.cues.map((cue) => cue.endTime)));
  }

  if (dropped.size === 0) {
    return cues;
  }
  return cues.flatMap((cue, index) => {
    if (dropped.has(index)) {
      return [];
    }
    const endTime = extendedEnd.get(index);
    return endTime !== undefined ? [{ ...cue, endTime }] : [cue];
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
export function hasAssAnimationEvidence(run: readonly AnnotatedSubtitleCue[]): boolean {
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
  const exactDeduplicated = collapseExactDuplicates(cues);
  const phaseDeduplicated =
    format === 'ass' ? collapseAnimatedStylePhases(exactDeduplicated) : exactDeduplicated;
  return collapseAnimationBursts(phaseDeduplicated, format);
}
