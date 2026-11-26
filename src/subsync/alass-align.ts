import {
  applyTimingToSubtitleFile,
  extractSpansFromParsedSubtitle,
  type ParsedSubtitleFile,
  type SubtitleTimeSpan,
} from './subtitle-lines';

const UNIT_MS = 10;
const UNITS_PER_SEC = 1000 / UNIT_MS; // 100
const MAX_SINGLE_SHIFT_RADIUS_SEC = 600; // ±10 minutes around median delta
const MAX_CUTS = 4;
const START_AGREE_SEC = 0.3;
const BROADCAST_PAUSE_SEC = 15.0;
const MIN_PIECE_LINES = 10;
const PIECE_Z_MIN = 2.0;
const PIECES_TOGETHER_Z_MIN = 6.5;
const MAX_PIECES = 6;
const STARTS_Z_MIN = 6.5;
const STANDING_PARTS = 4;

export interface SyncCalibration {
  readonly zero: number;
  readonly full: number;
  readonly cutsFrom: number;
}

export const CALIBRATION_NOISY: SyncCalibration = {
  zero: 2.0,
  full: 6.0,
  cutsFrom: 2.5,
};

export const CALIBRATION_CLEAN: SyncCalibration = {
  zero: 5.0,
  full: 7.5,
  cutsFrom: 7.0,
};

export interface AlignmentStanding {
  /** Z-score of overlap coverage vs shifted baselines (20s..5m away). */
  above: number;
  /** Z-score margin over the runner-up shifted peak (catches periodic false peaks). */
  alone: number;
  /** Line-start agreement z-score across the episode (and weakest quarter). */
  starts: number;
}

export interface AlignmentResult {
  /** 0..1 calibrated confidence score. */
  confidence: number;
  /** True when the alignment clears the automatic-apply confidence bar. */
  confident: boolean;
  /** Primary/dominant shift in milliseconds (including any audio lead adjustment). */
  offsetMs: number;
  /** Framerate ratio applied (1.0 when unchanged). */
  ratio: number;
  /** Number of distinct timing segments (1 = uniform shift, >1 = cuts/splits). */
  segments: number;
  /** True when multi-piece broadcast cut alignment was used instead of single-shift. */
  cut: boolean;
  standing: AlignmentStanding;
  /** Per-dialogue-line shift in seconds (indexed by dialogueIndices order). */
  movesSec: number[];
  dialogueIndices: number[];
  rewrittenContent: string;
}

export type ReferenceAudioStatus = 'aligned' | 'misaligned' | 'unverified';

export interface ReferenceAudioEvaluation {
  status: ReferenceAudioStatus;
  /** Shift in ms that best aligns this reference to audio (0 = already aligned). */
  offsetMs: number;
  /** Calibrated confidence of the best shifted alignment to audio (0..1). */
  confidence: number;
  /** Calibrated confidence at offset = 0 (as-is alignment to audio). */
  zeroConfidence: number;
  standing: AlignmentStanding;
}

/**
 * Merge overlapping spans so stacked lines (such as multi-layer karaoke or
 * simultaneous typeset signs) only count once per second.
 */
export function joinSpans(spans: readonly SubtitleTimeSpan[]): SubtitleTimeSpan[] {
  const sorted = spans
    .filter((s) => Number.isFinite(s.start) && Number.isFinite(s.end) && s.end > s.start)
    .map((s) => ({ start: s.start, end: s.end }))
    .sort((a, b) => a.start - b.start || a.end - b.end);

  const out: SubtitleTimeSpan[] = [];
  for (const span of sorted) {
    const last = out.at(-1);
    if (last && span.start <= last.end) {
      if (span.end > last.end) last.end = span.end;
    } else {
      out.push(span);
    }
  }
  return out;
}

/** Total seconds of overlap between sorted `lines` shifted by `shift` and disjoint sorted `ref`. */
export function coveredAt(
  sortedLines: readonly SubtitleTimeSpan[],
  joinedRef: readonly SubtitleTimeSpan[],
  shift: number,
): number {
  let covered = 0;
  let j = 0;
  const refLen = joinedRef.length;
  for (let i = 0; i < sortedLines.length; i += 1) {
    const line = sortedLines[i]!;
    const s = line.start + shift;
    const e = line.end + shift;
    while (j < refLen && joinedRef[j]!.end <= s) {
      j += 1;
    }
    let k = j;
    while (k < refLen && joinedRef[k]!.start < e) {
      const ref = joinedRef[k]!;
      const overlap = Math.min(e, ref.end) - Math.max(s, ref.start);
      if (overlap > 0) covered += overlap;
      k += 1;
    }
  }
  return covered;
}

function medianMidpoint(spans: readonly SubtitleTimeSpan[]): number {
  if (spans.length === 0) return 0;
  const mids = spans.map((s) => (s.start + s.end) * 0.5).sort((a, b) => a - b);
  return mids[Math.floor(mids.length / 2)]!;
}

/**
 * Exact alass-style single-shift alignment using a 2nd-difference sweep over 10ms bins.
 * Each (ref_i, in_j) pair contributes a piecewise-linear trapezoid to overlap(shift),
 * whose second derivative has exactly 4 impulses. Two prefix sums recover the exact
 * overlap at every 10ms shift in O(N*M + bins) (~1-3ms for an episode).
 */
export function alignNoSplit(
  joinedRef: readonly SubtitleTimeSpan[],
  inSpans: readonly SubtitleTimeSpan[],
  options: { centerShiftSec?: number; radiusSec?: number } = {},
): number {
  if (joinedRef.length === 0 || inSpans.length === 0) return 0;

  const centerShiftSec =
    options.centerShiftSec ?? medianMidpoint(joinedRef) - medianMidpoint(inSpans);
  const radiusSec = options.radiusSec ?? MAX_SINGLE_SHIFT_RADIUS_SEC;

  const minUnit = Math.floor((centerShiftSec - radiusSec) * UNITS_PER_SEC);
  const maxUnit = Math.ceil((centerShiftSec + radiusSec) * UNITS_PER_SEC);
  const binCount = maxUnit - minUnit + 1;
  if (binCount <= 0) return 0;

  // Store 2nd differences in Int32Array (each impulse is ±1, max magnitude <= N*M << 2^31).
  const d2 = new Int32Array(binCount + 2);
  let initialSlope = 0;
  let initialValue = 0;

  const refStartUnits = new Int32Array(joinedRef.length);
  const refEndUnits = new Int32Array(joinedRef.length);
  for (let i = 0; i < joinedRef.length; i += 1) {
    const rs = Math.round(joinedRef[i]!.start * UNITS_PER_SEC);
    const re = Math.max(rs + 1, Math.round(joinedRef[i]!.end * UNITS_PER_SEC));
    refStartUnits[i] = rs;
    refEndUnits[i] = re;
  }

  const inStartUnits = new Int32Array(inSpans.length);
  const inEndUnits = new Int32Array(inSpans.length);
  for (let j = 0; j < inSpans.length; j += 1) {
    const is = Math.round(inSpans[j]!.start * UNITS_PER_SEC);
    const ie = Math.max(is + 1, Math.round(inSpans[j]!.end * UNITS_PER_SEC));
    inStartUnits[j] = is;
    inEndUnits[j] = ie;
  }

  const addImpulse = (u: number, delta: number): void => {
    if (u <= minUnit) {
      initialSlope += delta;
      initialValue += delta * (minUnit - u);
    } else if (u <= maxUnit) {
      d2[u - minUnit] = d2[u - minUnit]! + delta;
    }
  };

  for (let i = 0; i < joinedRef.length; i += 1) {
    const a = refStartUnits[i]!;
    const b = refEndUnits[i]!;
    for (let j = 0; j < inSpans.length; j += 1) {
      const c = inStartUnits[j]!;
      const d = inEndUnits[j]!;
      const u1 = a - d;
      const u4 = b - c;
      if (u4 <= minUnit || u1 >= maxUnit) continue;
      addImpulse(u1, 1);
      addImpulse(a - c, -1);
      addImpulse(b - d, -1);
      addImpulse(u4, 1);
    }
  }

  let slope = initialSlope;
  let value = initialValue;
  let bestVal = value;
  let bestIdx = 0;

  for (let k = 0; k < binCount; k += 1) {
    if (k > 0) {
      slope += d2[k]!;
      value += slope;
    }
    if (value > bestVal) {
      bestVal = value;
      bestIdx = k;
    }
  }

  return (minUnit + bestIdx) / UNITS_PER_SEC;
}

/**
 * Center the shift in the middle of the near-optimal overlap plateau (±0.5s),
 * since subtitle lines are slightly longer than spoken utterances.
 */
function centerShiftOnCoverage(
  sortedLines: readonly SubtitleTimeSpan[],
  joinedRef: readonly SubtitleTimeSpan[],
  baseShift: number,
): number {
  let total = 0;
  for (const l of sortedLines) total += Math.max(0, l.end - l.start);
  if (total <= 0) return baseShift;

  const scores = new Float64Array(101);
  let best = 0;
  let peak = 50;
  for (let i = 0; i <= 100; i += 1) {
    const delta = (i - 50) * 0.01;
    const score = coveredAt(sortedLines, joinedRef, baseShift + delta) / total;
    scores[i] = score;
    if (score > best) {
      best = score;
      peak = i;
    }
  }
  if (best <= 0) return baseShift;
  let lo = peak;
  let hi = peak;
  while (lo > 0 && scores[lo - 1]! >= best - 0.002) lo -= 1;
  while (hi < 100 && scores[hi + 1]! >= best - 0.002) hi += 1;
  return baseShift + ((lo + hi) * 0.5 - 50) * 0.01;
}

function lowerBound(arr: readonly number[], target: number): number {
  let lo = 0;
  let hi = arr.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (arr[mid]! < target) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function agreeingStartsCount(
  starts: readonly number[],
  sortedRefStarts: readonly number[],
  shift: number,
): number {
  let count = 0;
  const refLen = sortedRefStarts.length;
  for (let i = 0; i < starts.length; i += 1) {
    const target = starts[i]! + shift;
    const idx = lowerBound(sortedRefStarts, target - START_AGREE_SEC);
    if (idx < refLen && sortedRefStarts[idx]! <= target + START_AGREE_SEC) {
      count += 1;
    }
  }
  return count;
}

function centerShiftOnStarts(
  starts: readonly number[],
  sortedRefStarts: readonly number[],
  baseShift: number,
): number {
  let best = -1;
  let peak = 50;
  const counts = new Int32Array(101);
  for (let i = 0; i <= 100; i += 1) {
    const c = agreeingStartsCount(starts, sortedRefStarts, baseShift + (i - 50) * 0.01);
    counts[i] = c;
    if (c > best) {
      best = c;
      peak = i;
    }
  }
  let end = peak;
  while (end + 1 <= 100 && counts[end + 1] === best) end += 1;
  return baseShift + ((peak + end) * 0.5 - 50) * 0.01;
}

/**
 * How far the line-start agreement at `shift` stands above chance (z-score
 * against shifted offsets 23s to 197s away, wrapped around episode length).
 */
export function agreementZ(
  starts: readonly number[],
  sortedRefStarts: readonly number[],
  shift: number,
  lengthSec: number,
): number {
  if (starts.length === 0 || sortedRefStarts.length === 0 || lengthSec <= 0) return 0;

  const STEP = 0.05;
  const MID = 3950; // ±197.5s
  const last = 2 * MID;

  let here = 0;
  for (let i = -10; i <= 10; i += 1) {
    const c = agreeingStartsCount(starts, sortedRefStarts, shift + i * STEP);
    if (c > here) here = c;
  }

  const from = shift - MID * STEP;
  const round: number[] = [];
  for (let j = -2; j <= 2; j += 1) {
    const offset = j * lengthSec;
    for (let i = 0; i < sortedRefStarts.length; i += 1) {
      round.push(sortedRefStarts[i]! + offset);
    }
  }
  round.sort((a, b) => a - b);

  const change = new Int32Array(last + 2);
  const windowSpan = last * STEP;

  for (let i = 0; i < starts.length; i += 1) {
    const s = starts[i]!;
    let runA = -1;
    let runB = -1;
    const minR = s + from - START_AGREE_SEC;
    const maxR = s + from + windowSpan + START_AGREE_SEC;
    let idx = lowerBound(round, minR);
    while (idx < round.length && round[idx]! <= maxR) {
      const r = round[idx]!;
      const a = Math.max(0, Math.ceil((r - s - START_AGREE_SEC - from) / STEP));
      const b = Math.min(last, Math.floor((r - s + START_AGREE_SEC - from) / STEP));
      if (a <= b) {
        if (runA >= 0 && a <= runB + 1) {
          if (b > runB) runB = b;
        } else {
          if (runA >= 0) {
            change[runA] = change[runA]! + 1;
            change[runB + 1] = change[runB + 1]! - 1;
          }
          runA = a;
          runB = b;
        }
      }
      idx += 1;
    }
    if (runA >= 0) {
      change[runA] = change[runA]! + 1;
      change[runB + 1] = change[runB + 1]! - 1;
    }
  }

  const counts = new Int32Array(last + 1);
  let running = 0;
  for (let i = 0; i <= last; i += 1) {
    running += change[i]!;
    counts[i] = running;
  }

  const others: number[] = [];
  for (let k = 23; k <= 197; k += 1) {
    for (const m of [MID + 20 * k, MID - 20 * k]) {
      let localMax = 0;
      for (let t = m - 10; t <= m + 10; t += 1) {
        const v = counts[t]!;
        if (v > localMax) localMax = v;
      }
      others.push(localMax);
    }
  }

  const n = others.length;
  let sum = 0;
  for (let i = 0; i < n; i += 1) sum += others[i]!;
  const mean = sum / n;
  let sq = 0;
  for (let i = 0; i < n; i += 1) {
    const diff = others[i]! - mean;
    sq += diff * diff;
  }
  const spread = Math.sqrt(sq / n);
  if (spread <= 1e-9) return 0;
  return (here - mean) / spread;
}

function startsStanding(
  placedStarts: readonly number[],
  sortedRefStarts: readonly number[],
  lengthSec: number,
): number {
  if (placedStarts.length === 0 || sortedRefStarts.length === 0) return 0;
  const sorted = [...placedStarts].sort((a, b) => a - b);
  const whole = agreementZ(sorted, sortedRefStarts, 0, lengthSec);
  const partSize = Math.max(1, Math.ceil(sorted.length / STANDING_PARTS));
  let weakest = Infinity;
  for (let offset = 0; offset < sorted.length; offset += partSize) {
    const part = sorted.slice(offset, offset + partSize);
    const z = agreementZ(part, sortedRefStarts, 0, lengthSec);
    if (z < weakest) weakest = z;
  }
  return weakest >= PIECE_Z_MIN ? whole : Math.min(weakest, whole);
}

function wrappedCoverage(
  sortedLines: readonly SubtitleTimeSpan[],
  total: number,
  joinedRef: readonly SubtitleTimeSpan[],
  shift: number,
  lengthSec: number,
): number {
  let covered = 0;
  if (shift >= 0) {
    let cut = 0;
    while (cut < sortedLines.length && sortedLines[cut]!.start + shift < lengthSec) cut += 1;
    covered =
      coveredAt(sortedLines.slice(0, cut), joinedRef, shift) +
      coveredAt(sortedLines.slice(cut), joinedRef, shift - lengthSec);
  } else {
    let cut = 0;
    while (cut < sortedLines.length && sortedLines[cut]!.start + shift < 0) cut += 1;
    covered =
      coveredAt(sortedLines.slice(0, cut), joinedRef, shift + lengthSec) +
      coveredAt(sortedLines.slice(cut), joinedRef, shift);
  }
  return covered / total;
}

export function evaluateStanding(
  placedSpans: readonly SubtitleTimeSpan[],
  joinedRef: readonly SubtitleTimeSpan[],
  sortedRefStarts: readonly number[],
  lengthSec: number,
): AlignmentStanding {
  const sorted = [...placedSpans].filter((s) => s.end > s.start).sort((a, b) => a.start - b.start);
  let total = 0;
  let maxEnd = 0;
  for (const s of sorted) {
    total += s.end - s.start;
    if (s.end > maxEnd) maxEnd = s.end;
  }
  const refEnd = joinedRef.length > 0 ? joinedRef[joinedRef.length - 1]!.end : 0;
  const effectiveLength = Math.max(lengthSec, maxEnd, refEnd);
  if (total <= 0 || effectiveLength <= 0 || joinedRef.length === 0) {
    return { above: 0, alone: 0, starts: 0 };
  }

  const here = coveredAt(sorted, joinedRef, 0) / total;
  const others: number[] = [];
  for (let k = 40; k <= 600; k += 1) {
    const shift = k * 0.5;
    others.push(wrappedCoverage(sorted, total, joinedRef, shift, effectiveLength));
    others.push(wrappedCoverage(sorted, total, joinedRef, -shift, effectiveLength));
  }

  const n = others.length;
  let sum = 0;
  let nextBest = 0;
  for (let i = 0; i < n; i += 1) {
    const v = others[i]!;
    sum += v;
    if (v > nextBest) nextBest = v;
  }
  const mean = sum / n;
  let sq = 0;
  for (let i = 0; i < n; i += 1) {
    const diff = others[i]! - mean;
    sq += diff * diff;
  }
  const spread = Math.sqrt(sq / n);
  const above = spread <= 1e-9 ? 0 : (here - mean) / spread;
  const alone = spread <= 1e-9 ? 0 : (here - nextBest) / spread;
  const starts = startsStanding(
    sorted.map((s) => s.start),
    sortedRefStarts,
    effectiveLength,
  );

  return { above, alone, starts };
}

export function scoreStanding(z: AlignmentStanding, cal: SyncCalibration): number {
  const raw = Math.max(0, Math.min(1, (z.above - cal.zero) / (cal.full - cal.zero)));
  const covered = z.alone < 0.3 ? Math.min(raw, 0.45) : raw;
  const started =
    z.starts >= STARTS_Z_MIN
      ? Math.max(
          0,
          Math.min(
            1,
            (z.starts - CALIBRATION_CLEAN.zero) / (CALIBRATION_CLEAN.full - CALIBRATION_CLEAN.zero),
          ),
        )
      : 0;
  return Math.max(covered, started);
}

/**
 * Viterbi piecewise refinement for 1..MAX_CUTS small scene cuts (±12s) around
 * a strong single shift, matching telly's piecewise refinement.
 */
function piecewiseAdjustments(
  placedSorted: readonly SubtitleTimeSpan[],
  joinedRef: readonly SubtitleTimeSpan[],
): number[] {
  const BLOCK = 16;
  const REACH = 120; // ±12.0s in 0.1s steps
  const JUMP = 8.0;

  const blocks: SubtitleTimeSpan[][] = [];
  for (let i = 0; i < placedSorted.length; i += BLOCK) {
    blocks.push(placedSorted.slice(i, i + BLOCK));
  }
  if (blocks.length < 2) {
    return new Array(placedSorted.length).fill(0);
  }

  const states = 2 * REACH + 1;
  const shiftOf = (s: number) => (s - REACH) * 0.1;

  const gains: Float64Array[] = blocks.map((block) => {
    const row = new Float64Array(states);
    for (let s = 0; s < states; s += 1) {
      row[s] = coveredAt(block, joinedRef, shiftOf(s));
    }
    return row;
  });

  let best = new Float64Array(states);
  for (let s = 0; s < states; s += 1) {
    best[s] = gains[0]![s]! - (s === REACH ? 0 : JUMP);
  }

  const back: Int32Array[] = [];
  for (let b = 1; b < blocks.length; b += 1) {
    let topState = 0;
    let topVal = -Infinity;
    for (let s = 0; s < states; s += 1) {
      if (best[s]! > topVal) {
        topVal = best[s]!;
        topState = s;
      }
    }
    const next = new Float64Array(states);
    const from = new Int32Array(states);
    const row = gains[b]!;
    for (let s = 0; s < states; s += 1) {
      if (best[s]! >= topVal - JUMP) {
        next[s] = best[s]! + row[s]!;
        from[s] = s;
      } else {
        next[s] = topVal - JUMP + row[s]!;
        from[s] = topState;
      }
    }
    back.push(from);
    best = next;
  }

  let state = 0;
  let maxEnd = -Infinity;
  for (let s = 0; s < states; s += 1) {
    if (best[s]! > maxEnd) {
      maxEnd = best[s]!;
      state = s;
    }
  }

  const path = new Int32Array(blocks.length);
  path[blocks.length - 1] = state;
  for (let b = blocks.length - 1; b >= 1; b -= 1) {
    state = back[b - 1]![state]!;
    path[b - 1] = state;
  }

  const extra = new Array<number>(placedSorted.length);
  for (let b = 0; b < blocks.length; b += 1) {
    const val = shiftOf(path[b]!);
    const base = b * BLOCK;
    for (let k = 0; k < blocks[b]!.length; k += 1) {
      extra[base + k] = val;
    }
  }

  // Refine each cut boundary to the single best line within the two neighboring blocks.
  for (let b = 1; b < blocks.length; b += 1) {
    const p = shiftOf(path[b - 1]!);
    const q = shiftOf(path[b]!);
    if (Math.abs(p - q) < 0.05) continue;
    const lo = (b - 1) * BLOCK;
    const hi = Math.min((b + 1) * BLOCK, placedSorted.length);
    let bestCut = b * BLOCK;
    let bestFit = -Infinity;
    for (let cut = lo; cut <= hi; cut += 1) {
      let fit = 0;
      for (let i = lo; i < hi; i += 1) {
        fit += coveredAt([placedSorted[i]!], joinedRef, i < cut ? p : q);
      }
      if (fit > bestFit) {
        bestFit = fit;
        bestCut = cut;
      }
    }
    for (let i = lo; i < hi; i += 1) {
      extra[i] = i < bestCut ? p : q;
    }
  }

  return extra;
}

interface BroadcastPiece {
  first: number;
  end: number;
  shift: number;
  z: number | null;
}

/**
 * Handle TV broadcast subtitles with large commercial/ad-break gaps (30s to 180s)
 * where a single shift cannot fit the whole episode. Splits along long pauses
 * (>= 15s), aligns each act independently via exact `alignNoSplit`, merges weak
 * or redundant pieces, and verifies the whole result with line-start agreement z-scores.
 */
function alignBroadcastPieces(
  scaledSorted: readonly SubtitleTimeSpan[],
  joinedRef: readonly SubtitleTimeSpan[],
  sortedRefStarts: readonly number[],
  lengthSec: number,
): { shifts: number[]; count: number; main: number; standing: AlignmentStanding } | null {
  if (scaledSorted.length < MIN_PIECE_LINES * 2) return null;

  const starts = scaledSorted.map((s) => s.start);
  const effectiveLength = Math.max(lengthSec, scaledSorted[scaledSorted.length - 1]!.end, 1.0);

  // Split at broadcast pauses (>= BROADCAST_PAUSE_SEC between end of previous dialogue and next start).
  const boundaries: number[] = [0];
  let maxEnded = scaledSorted[0]!.end;
  for (let i = 1; i < scaledSorted.length; i += 1) {
    const gap = scaledSorted[i]!.start - maxEnded;
    if (gap >= BROADCAST_PAUSE_SEC) {
      boundaries.push(i);
    }
    if (scaledSorted[i]!.end > maxEnded) maxEnded = scaledSorted[i]!.end;
  }
  boundaries.push(scaledSorted.length);

  if (boundaries.length - 1 < 2) return null;

  const pieces: BroadcastPiece[] = [];
  for (let b = 0; b < boundaries.length - 1; b += 1) {
    const first = boundaries[b]!;
    const end = boundaries[b + 1]!;
    const slice = scaledSorted.slice(first, end);
    const coarse = alignNoSplit(joinedRef, slice, { radiusSec: MAX_SINGLE_SHIFT_RADIUS_SEC });
    const refined = centerShiftOnStarts(starts.slice(first, end), sortedRefStarts, coarse);
    pieces.push({ first, end, shift: refined, z: null });
  }

  // Merge adjacent pieces that already share essentially the same shift.
  let idx = 1;
  while (idx < pieces.length) {
    const prev = pieces[idx - 1]!;
    const curr = pieces[idx]!;
    if (Math.abs(prev.shift - curr.shift) <= 0.4) {
      const mergedShift = centerShiftOnStarts(
        starts.slice(prev.first, curr.end),
        sortedRefStarts,
        curr.end - curr.first > prev.end - prev.first ? curr.shift : prev.shift,
      );
      pieces[idx - 1] = { first: prev.first, end: curr.end, shift: mergedShift, z: null };
      pieces.splice(idx, 1);
    } else {
      idx += 1;
    }
  }

  // Merge pieces too small or no better than chance into their neighbor.
  while (pieces.length > 1) {
    let weakIdx = -1;
    for (let i = 0; i < pieces.length; i += 1) {
      const p = pieces[i]!;
      if (p.z === null) {
        p.z = agreementZ(starts.slice(p.first, p.end), sortedRefStarts, p.shift, effectiveLength);
      }
      if (p.end - p.first < MIN_PIECE_LINES || p.z < PIECE_Z_MIN) {
        weakIdx = i;
        break;
      }
    }
    if (weakIdx === -1) break;
    const neighborIdx = weakIdx === 0 ? 1 : weakIdx - 1;
    const lo = Math.min(weakIdx, neighborIdx);
    const hi = Math.max(weakIdx, neighborIdx);
    const a = pieces[lo]!;
    const b = pieces[hi]!;
    const baseShift = pieces[neighborIdx]!.shift;
    const mergedShift = centerShiftOnStarts(
      starts.slice(a.first, b.end),
      sortedRefStarts,
      baseShift,
    );
    pieces[lo] = { first: a.first, end: b.end, shift: mergedShift, z: null };
    pieces.splice(hi, 1);
  }

  if (pieces.length < 2 || pieces.length > MAX_PIECES) return null;

  let weightedZSum = 0;
  let totalLines = 0;
  let weakest = Infinity;
  for (const p of pieces) {
    if (p.z === null) {
      p.z = agreementZ(starts.slice(p.first, p.end), sortedRefStarts, p.shift, effectiveLength);
    }
    const count = p.end - p.first;
    weightedZSum += Math.sqrt(count) * p.z;
    totalLines += count;
    if (p.z < weakest) weakest = p.z;
  }

  const together = totalLines > 0 ? weightedZSum / Math.sqrt(totalLines) : 0;
  if (together < PIECES_TOGETHER_Z_MIN || weakest < PIECE_Z_MIN) {
    return null;
  }

  const shifts = new Array<number>(scaledSorted.length);
  let mainPiece = pieces[0]!;
  for (const p of pieces) {
    if (p.end - p.first > mainPiece.end - mainPiece.first) {
      mainPiece = p;
    }
    for (let i = p.first; i < p.end; i += 1) {
      shifts[i] = p.shift;
    }
  }

  return {
    shifts,
    count: pieces.length,
    main: mainPiece.shift,
    standing: { above: together, alone: weakest, starts: together },
  };
}

export interface AlignSubtitleOptions {
  /** Whether to strip signs/songs when computing alignment (default true). */
  stripSigns?: boolean;
  /** Whether to allow scene-cut / broadcast split alignment (default true). */
  allowSplits?: boolean;
  /** Confidence calibration profile ('clean' for subtitle ref or Silero VAD, 'noisy' for PCM VAD). */
  calibration?: 'clean' | 'noisy';
  /** Media duration in seconds (optional; falls back to max span end). */
  lengthSec?: number;
  /**
   * Seconds subtitles typically appear ahead of speech onset (0 for subtitle ref,
   * ~0.1 for audio VAD).
   */
  leadSec?: number;
  /** Confidence threshold required to consider the sync automatic-apply safe. */
  confidenceThreshold?: number;
}

const FRAMERATE_RATIOS = [
  1.0,
  25.0 / 24.0,
  25.0 / 23.976,
  24.0 / 25.0,
  24.0 / 23.976,
  23.976 / 25.0,
  23.976 / 24.0,
];

export function alignSubtitleFileToReferenceSpans(
  target: ParsedSubtitleFile,
  rawReferenceSpans: readonly SubtitleTimeSpan[],
  options: AlignSubtitleOptions = {},
): AlignmentResult {
  const stripSigns = options.stripSigns !== false;
  const allowSplits = options.allowSplits !== false;
  const cal = options.calibration === 'noisy' ? CALIBRATION_NOISY : CALIBRATION_CLEAN;
  const leadSec = options.leadSec ?? 0;
  const confidenceThreshold = options.confidenceThreshold ?? 0.65;

  const { spans: targetSpans, dialogueIndices } = extractSpansFromParsedSubtitle(target, {
    stripSigns,
  });

  const sortedRefStarts = rawReferenceSpans
    .filter((s) => Number.isFinite(s.start) && Number.isFinite(s.end) && s.end > s.start)
    .map((s) => s.start)
    .sort((a, b) => a - b);
  const joinedRef = joinSpans(rawReferenceSpans);

  if (targetSpans.length === 0 || joinedRef.length === 0) {
    return {
      confidence: 0,
      confident: false,
      offsetMs: 0,
      ratio: 1.0,
      segments: 0,
      cut: false,
      standing: { above: 0, alone: 0, starts: 0 },
      movesSec: [],
      dialogueIndices,
      rewrittenContent: target.rawLines.join(target.eol),
    };
  }

  const maxTargetEnd = targetSpans.reduce((acc, s) => Math.max(acc, s.end), 0);
  const lengthSec = Math.max(
    options.lengthSec ?? 0,
    maxTargetEnd,
    joinedRef[joinedRef.length - 1]!.end,
    1.0,
  );

  // 1. Evaluate 1.0 and standard PAL/film framerate ratios.
  const evaluateRatio = (
    ratio: number,
  ): { ratio: number; shift: number; standing: AlignmentStanding } => {
    const scaled = targetSpans.map((s) => ({ start: s.start * ratio, end: s.end * ratio }));
    const sortedScaled = [...scaled].sort((a, b) => a.start - b.start);
    const coarseShift = alignNoSplit(joinedRef, sortedScaled);
    const centeredShift = centerShiftOnCoverage(sortedScaled, joinedRef, coarseShift);
    const moved = sortedScaled.map((s) => ({
      start: s.start + centeredShift,
      end: s.end + centeredShift,
    }));
    const standing = evaluateStanding(moved, joinedRef, sortedRefStarts, lengthSec);
    return { ratio, shift: centeredShift, standing };
  };

  const plain = evaluateRatio(1.0);
  let chosen = plain;

  // Only test framerate conversions if plain fit isn't already near-perfect or to catch PAL speedup.
  const altFits = FRAMERATE_RATIOS.slice(1).map(evaluateRatio);
  const bestAltByAbove = altFits.reduce(
    (best, cur) => (cur.standing.above > best.standing.above ? cur : best),
    altFits[0]!,
  );
  const bestAltByStarts = altFits.reduce(
    (best, cur) => (cur.standing.starts > best.standing.starts ? cur : best),
    altFits[0]!,
  );

  if (
    bestAltByAbove.standing.above > plain.standing.above + 2.0 &&
    bestAltByAbove.standing.above >= Math.max(4.0, cal.cutsFrom)
  ) {
    chosen = bestAltByAbove;
  } else if (
    bestAltByStarts.standing.starts >= STARTS_Z_MIN &&
    bestAltByStarts.standing.starts > plain.standing.starts + 2.0
  ) {
    chosen = bestAltByStarts;
  }

  const ratio = chosen.ratio;
  const oneShift = chosen.shift;
  const zOne = chosen.standing;
  const oneScore = scoreStanding(zOne, cal);

  // Order indices by start time for piecewise / broadcast passes.
  const order = targetSpans
    .map((_, idx) => idx)
    .sort((a, b) => targetSpans[a]!.start - targetSpans[b]!.start);
  const scaledSorted = order.map((idx) => ({
    start: targetSpans[idx]!.start * ratio,
    end: targetSpans[idx]!.end * ratio,
  }));

  let finalStanding = zOne;
  let finalScore = oneScore;
  let finalSegments = 1;
  let finalCut = false;
  let finalMainShift = oneShift;
  let finalMoves = new Array<number>(targetSpans.length).fill(oneShift);

  if (allowSplits) {
    const fitsWithOneShift = zOne.above >= cal.cutsFrom || zOne.starts >= STARTS_Z_MIN;

    if (fitsWithOneShift) {
      // Refine small scene cuts via Viterbi block DP.
      const placedSorted = scaledSorted.map((s) => ({
        start: s.start + oneShift,
        end: s.end + oneShift,
      }));
      const extra = piecewiseAdjustments(placedSorted, joinedRef);
      let jumps = 0;
      for (let i = 1; i < extra.length; i += 1) {
        if (Math.abs(extra[i]! - extra[i - 1]!) > 0.05) jumps += 1;
      }

      if (jumps >= 1 && jumps <= MAX_CUTS) {
        const splitMoves = new Array<number>(targetSpans.length);
        const splitSpans: SubtitleTimeSpan[] = [];
        for (let rank = 0; rank < order.length; rank += 1) {
          const idx = order[rank]!;
          const shift = oneShift + extra[rank]!;
          splitMoves[idx] = shift;
          splitSpans.push({
            start: targetSpans[idx]!.start * ratio + shift,
            end: targetSpans[idx]!.end * ratio + shift,
          });
        }
        const zSplit = evaluateStanding(splitSpans, joinedRef, sortedRefStarts, lengthSec);
        if (zSplit.above > zOne.above + 0.3) {
          finalStanding = zSplit;
          finalScore = scoreStanding(zSplit, cal);
          finalSegments = jumps + 1;
          finalMoves = splitMoves;
        }
      }
    }

    // If not every quarter fits the single/piecewise shift (e.g. a TV broadcast
    // with 30-180s commercial breaks where one act fits a single shift but others
    // don't), attempt multi-piece broadcast alignment verified by start agreement.
    if (finalStanding.starts < STARTS_Z_MIN) {
      const broadcast = alignBroadcastPieces(scaledSorted, joinedRef, sortedRefStarts, lengthSec);
      if (broadcast && broadcast.standing.starts > finalStanding.starts + 2.0) {
        const broadcastScore = scoreStanding(broadcast.standing, CALIBRATION_CLEAN);
        const pieceMoves = new Array<number>(targetSpans.length);
        for (let rank = 0; rank < order.length; rank += 1) {
          pieceMoves[order[rank]!] = broadcast.shifts[rank]!;
        }
        finalStanding = broadcast.standing;
        finalScore = broadcastScore;
        finalSegments = broadcast.count;
        finalCut = true;
        finalMainShift = broadcast.main;
        finalMoves = pieceMoves;
      }
    }
  }

  // Apply subtitle speech lead adjustment (e.g., -0.1s when syncing to audio speech).
  const adjustedMoves = leadSec !== 0 ? finalMoves.map((m) => m - leadSec) : finalMoves;
  const adjustedMainShift = finalMainShift - leadSec;

  const rewrittenContent = applyTimingToSubtitleFile(target, dialogueIndices, adjustedMoves, ratio);

  return {
    confidence: finalScore,
    confident: finalScore >= confidenceThreshold,
    offsetMs: Math.round(adjustedMainShift * 1000),
    ratio,
    segments: finalSegments,
    cut: finalCut,
    standing: finalStanding,
    movesSec: adjustedMoves,
    dialogueIndices,
    rewrittenContent,
  };
}

/**
 * Evaluate whether a candidate reference subtitle track is actually aligned with
 * the video's audio speech spans at offset ≈ 0, or if the reference subtitle itself
 * is shifted/misaligned relative to the audio.
 */
export function evaluateReferenceAlignmentToAudio(
  refDialogueSpans: readonly SubtitleTimeSpan[],
  rawAudioSpans: readonly SubtitleTimeSpan[],
  options: {
    lengthSec?: number;
    audioCalibration?: 'clean' | 'noisy';
    maxAlignedOffsetMs?: number;
    audioLeadSec?: number;
  } = {},
): ReferenceAudioEvaluation {
  const cal = options.audioCalibration === 'clean' ? CALIBRATION_CLEAN : CALIBRATION_NOISY;
  const maxAlignedOffsetMs = options.maxAlignedOffsetMs ?? 550;
  const audioLeadSec = options.audioLeadSec ?? 0.1;

  const sortedRef = refDialogueSpans
    .filter((s) => Number.isFinite(s.start) && Number.isFinite(s.end) && s.end > s.start)
    .sort((a, b) => a.start - b.start);
  const joinedAudio = joinSpans(rawAudioSpans);
  const sortedAudioStarts = rawAudioSpans
    .filter((s) => Number.isFinite(s.start) && Number.isFinite(s.end) && s.end > s.start)
    .map((s) => s.start)
    .sort((a, b) => a - b);

  if (sortedRef.length < 15 || joinedAudio.length === 0) {
    return {
      status: 'unverified',
      offsetMs: 0,
      confidence: 0,
      zeroConfidence: 0,
      standing: { above: 0, alone: 0, starts: 0 },
    };
  }

  const lengthSec = Math.max(
    options.lengthSec ?? 0,
    sortedRef[sortedRef.length - 1]!.end,
    joinedAudio[joinedAudio.length - 1]!.end,
    1.0,
  );

  // Measure as-is alignment (accounting for natural ~0.1s subtitle lead before speech).
  const atZero = sortedRef.map((s) => ({
    start: s.start + audioLeadSec,
    end: s.end + audioLeadSec,
  }));
  const zeroStanding = evaluateStanding(atZero, joinedAudio, sortedAudioStarts, lengthSec);
  const zeroConfidence = scoreStanding(zeroStanding, cal);

  // Find best shift between this reference subtitle and the audio speech.
  const coarse = alignNoSplit(joinedAudio, sortedRef);
  const bestShift = centerShiftOnCoverage(sortedRef, joinedAudio, coarse);
  const bestMoved = sortedRef.map((s) => ({
    start: s.start + bestShift,
    end: s.end + bestShift,
  }));
  const bestStanding = evaluateStanding(bestMoved, joinedAudio, sortedAudioStarts, lengthSec);
  const bestConfidence = scoreStanding(bestStanding, cal);

  // Effective offset of the reference relative to audio (subtracting natural subtitle lead).
  const offsetMs = Math.round((bestShift - audioLeadSec) * 1000);

  if (bestConfidence >= 0.45) {
    if (Math.abs(offsetMs) <= maxAlignedOffsetMs) {
      return {
        status: 'aligned',
        offsetMs,
        confidence: bestConfidence,
        zeroConfidence,
        standing: bestStanding,
      };
    }
    return {
      status: 'misaligned',
      offsetMs,
      confidence: bestConfidence,
      zeroConfidence,
      standing: bestStanding,
    };
  }

  if (zeroConfidence >= 0.45) {
    return {
      status: 'aligned',
      offsetMs: 0,
      confidence: zeroConfidence,
      zeroConfidence,
      standing: zeroStanding,
    };
  }

  return {
    status: 'unverified',
    offsetMs,
    confidence: bestConfidence,
    zeroConfidence,
    standing: bestStanding,
  };
}

export interface SampledAudioWindow {
  startSec: number;
  durationSec: number;
  speechSpans: SubtitleTimeSpan[];
}

/**
 * Evaluate whether a reference subtitle track aligns with the video's audio
 * using several seeked windows (e.g. 4 × 60s) without reading the full video.
 * Shifted baselines wrap strictly inside each window's own duration so unsampled
 * regions never create artificial silence bias. Returns `'unverified'` whenever
 * the sample is inconclusive so callers can escalate to full-file audio.
 */
export function evaluateReferenceAlignmentToSampledAudio(
  refDialogueSpans: readonly SubtitleTimeSpan[],
  windows: readonly SampledAudioWindow[],
  options: {
    maxAlignedOffsetMs?: number;
    audioLeadSec?: number;
    searchRadiusSec?: number;
  } = {},
): ReferenceAudioEvaluation {
  const maxAlignedOffsetMs = options.maxAlignedOffsetMs ?? 550;
  const audioLeadSec = options.audioLeadSec ?? 0.1;
  const searchRadiusSec = options.searchRadiusSec ?? 12.0;

  interface LocalWindow {
    durationSec: number;
    refSpans: SubtitleTimeSpan[];
    speechSpans: SubtitleTimeSpan[];
    totalRefSec: number;
  }

  const localWindows: LocalWindow[] = [];
  let totalRefLines = 0;
  let totalRefDuration = 0;

  for (const win of windows) {
    if (win.durationSec <= 15 || win.speechSpans.length === 0) continue;
    const winEnd = win.startSec + win.durationSec;
    const localRef: SubtitleTimeSpan[] = [];
    let refDur = 0;

    for (const span of refDialogueSpans) {
      if (!Number.isFinite(span.start) || !Number.isFinite(span.end) || span.end <= span.start) {
        continue;
      }
      const mid = (span.start + span.end) * 0.5;
      if (mid < win.startSec || mid > winEnd) continue;
      const s = Math.max(0, span.start - win.startSec);
      const e = Math.min(win.durationSec, span.end - win.startSec);
      if (e > s) {
        localRef.push({ start: s, end: e });
        refDur += e - s;
      }
    }

    if (localRef.length < 3 || refDur <= 0) continue;
    localRef.sort((a, b) => a.start - b.start);
    const localSpeech = joinSpans(
      win.speechSpans
        .map((s) => ({
          start: Math.max(0, s.start),
          end: Math.min(win.durationSec, s.end),
        }))
        .filter((s) => s.end > s.start),
    );
    if (localSpeech.length === 0) continue;

    localWindows.push({
      durationSec: win.durationSec,
      refSpans: localRef,
      speechSpans: localSpeech,
      totalRefSec: refDur,
    });
    totalRefLines += localRef.length;
    totalRefDuration += refDur;
  }

  if (localWindows.length < 2 || totalRefLines < 20 || totalRefDuration <= 0) {
    return {
      status: 'unverified',
      offsetMs: 0,
      confidence: 0,
      zeroConfidence: 0,
      standing: { above: 0, alone: 0, starts: 0 },
    };
  }

  const combinedCoverageAt = (shift: number): number => {
    let covered = 0;
    for (const w of localWindows) {
      covered += coveredAt(w.refSpans, w.speechSpans, shift);
    }
    return covered / totalRefDuration;
  };

  // Coarse 0.1s scan over [-searchRadiusSec, +searchRadiusSec], then fine 0.01s around peak.
  let bestShift = 0;
  let bestCov = -1;
  const coarseSteps = Math.round(searchRadiusSec * 10);
  for (let k = -coarseSteps; k <= coarseSteps; k += 1) {
    const s = k * 0.1;
    const cov = combinedCoverageAt(s);
    if (cov > bestCov) {
      bestCov = cov;
      bestShift = s;
    }
  }
  const coarseBest = bestShift;
  for (let k = -15; k <= 15; k += 1) {
    const s = coarseBest + k * 0.01;
    const cov = combinedCoverageAt(s);
    if (cov > bestCov) {
      bestCov = cov;
      bestShift = s;
    }
  }

  // Center within near-optimal plateau (±0.4s).
  const plateauScores: number[] = [];
  let peakVal = 0;
  let peakIdx = 40;
  for (let i = 0; i <= 80; i += 1) {
    const cov = combinedCoverageAt(bestShift + (i - 40) * 0.01);
    plateauScores.push(cov);
    if (cov > peakVal) {
      peakVal = cov;
      peakIdx = i;
    }
  }
  let lo = peakIdx;
  let hi = peakIdx;
  while (lo > 0 && plateauScores[lo - 1]! >= peakVal - 0.002) lo -= 1;
  while (hi < 80 && plateauScores[hi + 1]! >= peakVal - 0.002) hi += 1;
  bestShift += ((lo + hi) * 0.5 - 40) * 0.01;

  // Evaluate shifted null baselines strictly wrapped inside each window (±5s..±28s).
  const computeWrappedStanding = (targetShift: number): AlignmentStanding => {
    const shiftedWindows = localWindows.map((w) => ({
      durationSec: w.durationSec,
      speechSpans: w.speechSpans,
      refSpans: w.refSpans
        .map((s) => ({
          start: (((s.start + targetShift) % w.durationSec) + w.durationSec) % w.durationSec,
          end:
            ((((s.start + targetShift) % w.durationSec) + w.durationSec) % w.durationSec) +
            (s.end - s.start),
        }))
        .sort((a, b) => a.start - b.start),
    }));

    let hereCovered = 0;
    for (const w of shiftedWindows) {
      hereCovered += coveredAt(w.refSpans, w.speechSpans, 0);
    }
    const here = hereCovered / totalRefDuration;

    const others: number[] = [];
    for (let k = 10; k <= 56; k += 1) {
      const delta = k * 0.5;
      for (const dir of [delta, -delta]) {
        let cov = 0;
        for (const w of shiftedWindows) {
          cov += wrappedCoverage(w.refSpans, 1, w.speechSpans, dir, w.durationSec);
        }
        others.push(cov / totalRefDuration);
      }
    }

    const n = others.length;
    let sum = 0;
    let nextBest = 0;
    for (let i = 0; i < n; i += 1) {
      const v = others[i]!;
      sum += v;
      if (v > nextBest) nextBest = v;
    }
    const mean = sum / n;
    let sq = 0;
    for (let i = 0; i < n; i += 1) {
      const diff = others[i]! - mean;
      sq += diff * diff;
    }
    const spread = Math.sqrt(sq / n);
    if (spread <= 1e-9) return { above: 0, alone: 0, starts: 0 };
    return {
      above: (here - mean) / spread,
      alone: (here - nextBest) / spread,
      starts: 0,
    };
  };

  const bestStanding = computeWrappedStanding(bestShift);
  const zeroStanding = computeWrappedStanding(audioLeadSec);
  const bestConfidence = scoreStanding(bestStanding, CALIBRATION_NOISY);
  const zeroConfidence = scoreStanding(zeroStanding, CALIBRATION_NOISY);
  const offsetMs = Math.round((bestShift - audioLeadSec) * 1000);

  if (bestConfidence >= 0.5) {
    if (Math.abs(offsetMs) <= maxAlignedOffsetMs) {
      return {
        status: 'aligned',
        offsetMs,
        confidence: bestConfidence,
        zeroConfidence,
        standing: bestStanding,
      };
    }
    return {
      status: 'misaligned',
      offsetMs,
      confidence: bestConfidence,
      zeroConfidence,
      standing: bestStanding,
    };
  }

  return {
    status: 'unverified',
    offsetMs,
    confidence: bestConfidence,
    zeroConfidence,
    standing: bestStanding,
  };
}
