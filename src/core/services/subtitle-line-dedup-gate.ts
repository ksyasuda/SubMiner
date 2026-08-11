/*
 * Decides which live mpv subtitle lines reach the immersion stats.
 *
 * The sidebar reads a parsed subtitle file, so it can collapse an animation burst with
 * full lookahead (`subtitle-cue-dedup`). Stats are fed from mpv's `sub-start`/`sub-end`
 * properties instead -- one event per animation frame, each with its own start time --
 * so without a gate a karaoke OP counts its lyrics once per frame and buries every real
 * word in the vocabulary charts.
 *
 * Two layers, in order:
 *
 * 1. When the active source has been parsed, its cue list has *already* been collapsed.
 *    A live line that lands inside a surviving cue of the same text, but after that
 *    cue's start, is a frame the sidebar merged away, so stats drop it too. This is the
 *    layer that keeps the two views consistent by construction.
 * 2. Otherwise (embedded track nobody parsed, a source whose timings mpv has shifted)
 *    fall back to timing alone. No authoring metadata is available live -- mpv delivers
 *    `sub-text-ass` after `sub-start`/`sub-end`, so any ASS text read here belongs to the
 *    previous event -- which puts this layer in the same position as the SRT path in
 *    `subtitle-cue-dedup`, and it uses that path's deliberately strict bounds.
 */

import { normalizePlainSubtitleText } from './ass-text';
import {
  DUPLICATE_CUE_GAP_TOLERANCE_SECONDS,
  MIN_TIMING_ONLY_FRAMES,
  TIMING_ONLY_FRAME_MAX_SECONDS,
} from './subtitle-burst-constants';
import type { SubtitleCue } from './subtitle-cue-parser';

export interface SubtitleLineSample {
  text: string;
  startSec: number;
  endSec: number;
}

export interface SubtitleLineDedupGateDeps {
  /** Cues for the active source, already collapsed by the parser. */
  getParsedCues: () => readonly SubtitleCue[] | null | undefined;
}

export interface SubtitleLineDedupGate {
  /** False when this line is an animation frame of a line already recorded. */
  shouldRecord: (sample: SubtitleLineSample) => boolean;
  /** Forget the streaming run state, e.g. when playback moves to another file. */
  reset: () => void;
}

interface CueSpan {
  startTime: number;
  endTime: number;
}

interface StreamingRunState {
  text: string;
  startMs: number;
  chainEndSec: number;
  /** Contiguous identical short frames seen so far, including the recorded first one. */
  frames: number;
}

function normalizeLineText(text: string): string {
  return normalizePlainSubtitleText(text, { collapseLineBreaks: true });
}

function buildSpansByText(cues: readonly SubtitleCue[]): Map<string, CueSpan[]> {
  const spansByText = new Map<string, CueSpan[]>();
  for (const cue of cues) {
    const key = normalizeLineText(cue.text);
    if (!key) continue;
    const span = { startTime: cue.startTime, endTime: cue.endTime };
    const existing = spansByText.get(key);
    if (existing) {
      existing.push(span);
    } else {
      spansByText.set(key, [span]);
    }
  }
  return spansByText;
}

/**
 * A frame the parser merged away: the same text, starting inside a surviving cue but
 * after it began.
 *
 * Starting a cue always wins over falling inside one. The first frame of a collapsed run
 * starts *at* the merged cue, and a line the parser deliberately kept separate -- three
 * characters trading `えっ` back to back -- begins exactly where the one before it ends.
 */
function isMergedAwayFrame(spans: readonly CueSpan[], startSec: number): boolean {
  const startsOwnCue = spans.some(
    (span) => Math.abs(startSec - span.startTime) <= DUPLICATE_CUE_GAP_TOLERANCE_SECONDS,
  );
  if (startsOwnCue) {
    return false;
  }
  return spans.some(
    (span) =>
      startSec > span.startTime + DUPLICATE_CUE_GAP_TOLERANCE_SECONDS &&
      startSec <= span.endTime + DUPLICATE_CUE_GAP_TOLERANCE_SECONDS,
  );
}

export function createSubtitleLineDedupGate(
  deps: SubtitleLineDedupGateDeps,
): SubtitleLineDedupGate {
  let indexedCues: readonly SubtitleCue[] | null = null;
  let spansByText: Map<string, CueSpan[]> = new Map();
  let run: StreamingRunState | null = null;

  const lookupSpans = (text: string): CueSpan[] | null => {
    const cues = deps.getParsedCues();
    if (!cues?.length) {
      indexedCues = null;
      spansByText = new Map();
      return null;
    }
    if (cues !== indexedCues) {
      indexedCues = cues;
      spansByText = buildSpansByText(cues);
    }
    return spansByText.get(text) ?? null;
  };

  /**
   * Timing-only burst detection over a stream. Without lookahead the run can only be
   * recognised from the inside, so the first frames of a burst are recorded and the rest
   * dropped -- an OP costs a handful of counted lines instead of several hundred.
   */
  const advanceStreamingRun = (text: string, sample: SubtitleLineSample): boolean => {
    const startMs = Math.round(sample.startSec * 1000);
    // mpv reports `sub-start` and `sub-end` separately, so one event can be offered
    // twice. The same start is the same frame, never the next one in a run.
    if (run && run.text === text && run.startMs === startMs) {
      run.chainEndSec = Math.max(run.chainEndSec, sample.endSec);
      return run.frames < MIN_TIMING_ONLY_FRAMES;
    }

    const isShortFrame = sample.endSec - sample.startSec < TIMING_ONLY_FRAME_MAX_SECONDS;
    // Frames are authored flush against each other, but typesetters do overlap them, so
    // the chain only requires forward progress that stays inside the running end.
    const continuesRun =
      run !== null &&
      run.text === text &&
      isShortFrame &&
      startMs > run.startMs &&
      sample.startSec <= run.chainEndSec + DUPLICATE_CUE_GAP_TOLERANCE_SECONDS;

    if (continuesRun && run) {
      run.startMs = startMs;
      run.chainEndSec = Math.max(run.chainEndSec, sample.endSec);
      run.frames += 1;
    } else {
      run = {
        text,
        startMs,
        chainEndSec: sample.endSec,
        frames: isShortFrame ? 1 : 0,
      };
    }

    return run.frames < MIN_TIMING_ONLY_FRAMES;
  };

  return {
    shouldRecord: (sample) => {
      const text = normalizeLineText(sample.text);
      if (!text) {
        return true;
      }

      const spans = lookupSpans(text);
      if (spans && isMergedAwayFrame(spans, sample.startSec)) {
        run = null;
        return false;
      }

      return advanceStreamingRun(text, sample);
    },
    reset: () => {
      run = null;
    },
  };
}
