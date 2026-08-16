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
 *    `sub-text/ass` after `sub-start`/`sub-end`, so any ASS text read here belongs to the
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
  /** Forget run state and ignore the current cue list until its source is replaced. */
  reset: () => void;
}

interface CueSpan {
  startTime: number;
  endTime: number;
}

interface StreamingRunState {
  startMs: number;
  chainEndSec: number;
  /** Contiguous identical short frames seen so far, including the recorded first one. */
  frames: number;
}

/**
 * Dual-line karaoke interleaves two texts frame by frame, so runs are tracked per text.
 * Dead runs are pruned as playback moves past them; the cap only matters after a
 * backward seek leaves runs whose ends sit ahead of the new position.
 */
const MAX_ACTIVE_STREAMING_RUNS = 32;

/** Exact cue identity, separate from the looser tolerance used to chain adjacent frames. */
const CUE_START_IDENTITY_TOLERANCE_SECONDS = 0.005;

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
function isMergedAwayFrame(spans: readonly CueSpan[], startSec: number): boolean | null {
  const coveringSpans = spans.filter(
    (span) =>
      startSec >= span.startTime - CUE_START_IDENTITY_TOLERANCE_SECONDS &&
      startSec <= span.endTime + CUE_START_IDENTITY_TOLERANCE_SECONDS,
  );
  if (coveringSpans.length === 0) {
    return null;
  }
  const startsOwnCue = spans.some(
    (span) => Math.abs(startSec - span.startTime) <= CUE_START_IDENTITY_TOLERANCE_SECONDS,
  );
  if (startsOwnCue) {
    return false;
  }
  return coveringSpans.some(
    (span) =>
      startSec > span.startTime + CUE_START_IDENTITY_TOLERANCE_SECONDS &&
      startSec <= span.endTime + CUE_START_IDENTITY_TOLERANCE_SECONDS,
  );
}

export function createSubtitleLineDedupGate(
  deps: SubtitleLineDedupGateDeps,
): SubtitleLineDedupGate {
  let indexedCues: readonly SubtitleCue[] | null | undefined;
  let ignoredCuesAfterReset: readonly SubtitleCue[] | null | undefined;
  let spansByText: Map<string, CueSpan[]> = new Map();
  const runs = new Map<string, StreamingRunState>();

  const lookupSpans = (text: string): CueSpan[] | null => {
    const cues = deps.getParsedCues() ?? null;
    if (ignoredCuesAfterReset !== undefined) {
      if (cues === ignoredCuesAfterReset) {
        return null;
      }
      ignoredCuesAfterReset = undefined;
    }
    if (cues !== indexedCues) {
      indexedCues = cues;
      spansByText = cues?.length ? buildSpansByText(cues) : new Map();
      runs.clear();
    }
    return spansByText.get(text) ?? null;
  };

  /**
   * A run this sample cannot continue is a run no later sample can continue either --
   * continuation needs a start inside the running end plus tolerance, and starts only
   * move forward outside of seeks.
   */
  const pruneDeadRuns = (startSec: number): void => {
    for (const [text, state] of runs) {
      if (state.chainEndSec + DUPLICATE_CUE_GAP_TOLERANCE_SECONDS < startSec) {
        runs.delete(text);
      }
    }
  };

  /**
   * Timing-only burst detection over a stream. Without lookahead the run can only be
   * recognised from the inside, so the first frames of a burst are recorded and the rest
   * dropped -- an OP costs a handful of counted lines instead of several hundred.
   */
  const advanceStreamingRun = (text: string, sample: SubtitleLineSample): boolean => {
    pruneDeadRuns(sample.startSec);
    const startMs = Math.round(sample.startSec * 1000);
    const run = runs.get(text);
    // mpv reports `sub-start` and `sub-end` separately, so one event can be offered
    // twice. The same start is the same frame, never the next one in a run.
    if (run && run.startMs === startMs) {
      run.chainEndSec = Math.max(run.chainEndSec, sample.endSec);
      return run.frames < MIN_TIMING_ONLY_FRAMES;
    }

    const isShortFrame = sample.endSec - sample.startSec < TIMING_ONLY_FRAME_MAX_SECONDS;
    // Frames are authored flush against each other, but typesetters do overlap them, so
    // the chain only requires forward progress that stays inside the running end.
    const continuesRun =
      run !== undefined &&
      isShortFrame &&
      startMs > run.startMs &&
      sample.startSec <= run.chainEndSec + DUPLICATE_CUE_GAP_TOLERANCE_SECONDS;

    if (continuesRun && run) {
      run.startMs = startMs;
      run.chainEndSec = Math.max(run.chainEndSec, sample.endSec);
      run.frames += 1;
      return run.frames < MIN_TIMING_ONLY_FRAMES;
    }

    const fresh: StreamingRunState = {
      startMs,
      chainEndSec: sample.endSec,
      frames: isShortFrame ? 1 : 0,
    };
    runs.set(text, fresh);
    if (runs.size > MAX_ACTIVE_STREAMING_RUNS) {
      let oldestText: string | undefined;
      let oldestEnd = Infinity;
      for (const [runText, state] of runs) {
        if (runText !== text && state.chainEndSec < oldestEnd) {
          oldestEnd = state.chainEndSec;
          oldestText = runText;
        }
      }
      if (oldestText !== undefined) {
        runs.delete(oldestText);
      }
    }
    return fresh.frames < MIN_TIMING_ONLY_FRAMES;
  };

  return {
    shouldRecord: (sample) => {
      const text = normalizeLineText(sample.text);
      if (!text) {
        return true;
      }

      // The parsed cue list has the final say wherever it covers this line. Falling
      // through to the streaming heuristic would let it drop cues the parser looked at
      // with full lookahead and deliberately kept apart, which is the disagreement
      // between sidebar and stats this gate exists to prevent.
      const spans = lookupSpans(text);
      if (spans) {
        const mergedAway = isMergedAwayFrame(spans, sample.startSec);
        if (mergedAway !== null) {
          runs.delete(text);
          return !mergedAway;
        }
      }

      return advanceStreamingRun(text, sample);
    },
    reset: () => {
      runs.clear();
      ignoredCuesAfterReset = deps.getParsedCues() ?? null;
      indexedCues = undefined;
      spansByText = new Map();
    },
  };
}
