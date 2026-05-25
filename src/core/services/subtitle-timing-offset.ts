import type { SubtitleCue } from './subtitle-cue-parser';

export type SubtitleTimingOffsetResult = {
  offsetSeconds: number;
  matchCount: number;
  meanErrorSeconds: number;
  maxErrorSeconds: number;
};

export type SubtitleTimingOffsetOptions = {
  maxCueCount?: number;
  maxOffsetSeconds?: number;
  matchThresholdSeconds?: number;
  maxMeanErrorSeconds?: number;
  minMatchCount?: number;
  minMatchRatio?: number;
  minUsefulOffsetSeconds?: number;
};

type OffsetScore = SubtitleTimingOffsetResult;

const DEFAULT_MAX_CUE_COUNT = 60;
const DEFAULT_MAX_OFFSET_SECONDS = 180;
const DEFAULT_MATCH_THRESHOLD_SECONDS = 1;
const DEFAULT_MAX_MEAN_ERROR_SECONDS = 0.75;
const DEFAULT_MIN_MATCH_COUNT = 8;
const DEFAULT_MIN_MATCH_RATIO = 0.25;
const DEFAULT_MIN_USEFUL_OFFSET_SECONDS = 0.25;

function normalizeCueStarts(cues: SubtitleCue[], maxCueCount: number): number[] {
  const starts = cues
    .map((cue) => cue.startTime)
    .filter((start) => Number.isFinite(start) && start >= 0)
    .sort((a, b) => a - b);
  const deduped: number[] = [];
  for (const start of starts) {
    const previous = deduped[deduped.length - 1];
    if (previous === undefined || Math.abs(start - previous) > 0.05) {
      deduped.push(start);
    }
    if (deduped.length >= maxCueCount) {
      break;
    }
  }
  return deduped;
}

function roundToMillis(value: number): number {
  return Math.round(value * 1000) / 1000;
}

function scoreOffset(
  primaryStarts: number[],
  referenceStarts: number[],
  offsetSeconds: number,
  matchThresholdSeconds: number,
): OffsetScore {
  let primaryIndex = 0;
  let referenceIndex = 0;
  let matchCount = 0;
  let totalErrorSeconds = 0;
  let maxErrorSeconds = 0;

  while (primaryIndex < primaryStarts.length && referenceIndex < referenceStarts.length) {
    const shiftedPrimary = primaryStarts[primaryIndex]! + offsetSeconds;
    const reference = referenceStarts[referenceIndex]!;
    const errorSeconds = Math.abs(shiftedPrimary - reference);
    if (errorSeconds <= matchThresholdSeconds) {
      matchCount += 1;
      totalErrorSeconds += errorSeconds;
      maxErrorSeconds = Math.max(maxErrorSeconds, errorSeconds);
      primaryIndex += 1;
      referenceIndex += 1;
      continue;
    }

    if (shiftedPrimary < reference) {
      primaryIndex += 1;
    } else {
      referenceIndex += 1;
    }
  }

  return {
    offsetSeconds,
    matchCount,
    meanErrorSeconds: matchCount > 0 ? totalErrorSeconds / matchCount : Number.POSITIVE_INFINITY,
    maxErrorSeconds,
  };
}

function isBetterScore(next: OffsetScore, current: OffsetScore | null): boolean {
  if (current === null) return true;
  if (next.matchCount !== current.matchCount) return next.matchCount > current.matchCount;
  if (next.meanErrorSeconds !== current.meanErrorSeconds) {
    return next.meanErrorSeconds < current.meanErrorSeconds;
  }
  return Math.abs(next.offsetSeconds) < Math.abs(current.offsetSeconds);
}

export function estimateSubtitleTimingOffset(
  primaryCues: SubtitleCue[],
  referenceCues: SubtitleCue[],
  options: SubtitleTimingOffsetOptions = {},
): SubtitleTimingOffsetResult | null {
  const maxCueCount = options.maxCueCount ?? DEFAULT_MAX_CUE_COUNT;
  const maxOffsetSeconds = options.maxOffsetSeconds ?? DEFAULT_MAX_OFFSET_SECONDS;
  const matchThresholdSeconds = options.matchThresholdSeconds ?? DEFAULT_MATCH_THRESHOLD_SECONDS;
  const maxMeanErrorSeconds = options.maxMeanErrorSeconds ?? DEFAULT_MAX_MEAN_ERROR_SECONDS;
  const minMatchCount = options.minMatchCount ?? DEFAULT_MIN_MATCH_COUNT;
  const minMatchRatio = options.minMatchRatio ?? DEFAULT_MIN_MATCH_RATIO;
  const minUsefulOffsetSeconds =
    options.minUsefulOffsetSeconds ?? DEFAULT_MIN_USEFUL_OFFSET_SECONDS;

  const primaryStarts = normalizeCueStarts(primaryCues, maxCueCount);
  const referenceStarts = normalizeCueStarts(referenceCues, maxCueCount);
  const comparableCueCount = Math.min(primaryStarts.length, referenceStarts.length);
  if (comparableCueCount < minMatchCount) {
    return null;
  }

  const candidates = new Set<number>();
  for (const primaryStart of primaryStarts) {
    for (const referenceStart of referenceStarts) {
      const offsetSeconds = roundToMillis(referenceStart - primaryStart);
      if (Math.abs(offsetSeconds) <= maxOffsetSeconds) {
        candidates.add(offsetSeconds);
      }
    }
  }

  let best: OffsetScore | null = null;
  for (const offsetSeconds of candidates) {
    if (Math.abs(offsetSeconds) < minUsefulOffsetSeconds) {
      continue;
    }
    const score = scoreOffset(primaryStarts, referenceStarts, offsetSeconds, matchThresholdSeconds);
    if (score.matchCount < minMatchCount) {
      continue;
    }
    if (score.matchCount / comparableCueCount < minMatchRatio) {
      continue;
    }
    if (score.meanErrorSeconds > maxMeanErrorSeconds) {
      continue;
    }
    if (isBetterScore(score, best)) {
      best = score;
    }
  }

  return best;
}
