import type { KnownWordsTimelinePoint } from '../hooks/useSessions';
import { EventType } from '../types/stats';
import type { SessionEvent, SessionTimelinePoint } from '../types/stats';

/** One time slice of a session: words read, lookups made, and cards mined inside it. */
export interface SessionDensityBucket {
  index: number;
  startMs: number;
  endMs: number;
  words: number;
  knownWords: number;
  unknownWords: number;
  lookups: number;
  cards: number;
}

export interface SessionDensity {
  bucketMs: number;
  startMs: number;
  endMs: number;
  /** True when the known-words timeline had data, so known/unknown splits are meaningful. */
  hasKnownWords: boolean;
  buckets: SessionDensityBucket[];
}

const BUCKET_SIZE_CANDIDATES_MS = [10_000, 15_000, 30_000, 60_000, 120_000, 300_000, 600_000];
const LARGEST_BUCKET_MS = 1_800_000;

export const MAX_DENSITY_BUCKETS = 60;

/** Smallest candidate bucket size that keeps the whole span within MAX_DENSITY_BUCKETS. */
export function pickDensityBucketMs(spanMs: number, maxBuckets = MAX_DENSITY_BUCKETS): number {
  for (const candidate of BUCKET_SIZE_CANDIDATES_MS) {
    if (Math.ceil(spanMs / candidate) <= maxBuckets) return candidate;
  }
  return LARGEST_BUCKET_MS;
}

export function formatBucketSize(bucketMs: number): string {
  if (bucketMs < 60_000) return `${Math.round(bucketMs / 1000)}s`;
  const minutes = Math.round(bucketMs / 60_000);
  return `${minutes}m`;
}

type KnownWordCounts = { knownWordsSeen: number; totalWordsSeen: number };

/** Cumulative known/total counts at a lines-seen checkpoint, using the nearest earlier point. */
function lookupKnownWordCounts(
  sortedPoints: KnownWordsTimelinePoint[],
  linesSeen: number,
): KnownWordCounts {
  let best: KnownWordCounts = { knownWordsSeen: 0, totalWordsSeen: 0 };
  for (const point of sortedPoints) {
    if (point.linesSeen > linesSeen) break;
    best = point;
  }
  return best;
}

/**
 * Turn cumulative telemetry samples into per-bucket deltas so the chart shows how
 * dense the dialogue was over the session instead of a monotonically rising total.
 */
export function buildSessionDensity(
  timeline: SessionTimelinePoint[],
  knownWordsTimeline: KnownWordsTimelinePoint[],
  events: SessionEvent[],
): SessionDensity {
  const samples = [...timeline].sort((a, b) => a.sampleMs - b.sampleMs);
  if (samples.length === 0) {
    return { bucketMs: 0, startMs: 0, endMs: 0, hasKnownWords: false, buckets: [] };
  }

  const knownPoints = [...knownWordsTimeline].sort((a, b) => a.linesSeen - b.linesSeen);
  const hasKnownWords = knownPoints.length > 0;
  const startMs = samples[0]!.sampleMs;
  const lastSampleMs = samples[samples.length - 1]!.sampleMs;
  const spanMs = Math.max(1, lastSampleMs - startMs);
  const bucketMs = pickDensityBucketMs(spanMs);
  const bucketCount = Math.max(1, Math.ceil(spanMs / bucketMs));
  const endMs = startMs + bucketCount * bucketMs;

  const buckets: SessionDensityBucket[] = [];
  let sampleIndex = -1;
  let prevWords = 0;
  let prevKnown = 0;

  for (let index = 0; index < bucketCount; index += 1) {
    const bucketStart = startMs + index * bucketMs;
    const bucketEnd = bucketStart + bucketMs;
    // The last bucket also owns a sample sitting exactly on the session's end edge.
    const isLastBucket = index === bucketCount - 1;
    while (
      sampleIndex + 1 < samples.length &&
      (isLastBucket || samples[sampleIndex + 1]!.sampleMs < bucketEnd)
    ) {
      sampleIndex += 1;
    }

    let cumulativeWords = prevWords;
    let cumulativeKnown = prevKnown;
    if (sampleIndex >= 0) {
      const sample = samples[sampleIndex]!;
      if (hasKnownWords) {
        const counts = lookupKnownWordCounts(knownPoints, sample.linesSeen);
        cumulativeWords = counts.totalWordsSeen;
        cumulativeKnown = Math.min(counts.knownWordsSeen, counts.totalWordsSeen);
      } else {
        cumulativeWords = sample.tokensSeen;
      }
    }

    const words = Math.max(0, cumulativeWords - prevWords);
    const knownWords = Math.min(words, Math.max(0, cumulativeKnown - prevKnown));
    prevWords = Math.max(prevWords, cumulativeWords);
    prevKnown = Math.max(prevKnown, cumulativeKnown);

    buckets.push({
      index,
      startMs: bucketStart,
      endMs: bucketEnd,
      words,
      knownWords,
      unknownWords: words - knownWords,
      lookups: 0,
      cards: 0,
    });
  }

  for (const event of events) {
    if (event.eventType !== EventType.YOMITAN_LOOKUP && event.eventType !== EventType.CARD_MINED) {
      continue;
    }
    const rawIndex = Math.floor((event.tsMs - startMs) / bucketMs);
    const bucket = buckets[Math.max(0, Math.min(bucketCount - 1, rawIndex))]!;
    if (event.eventType === EventType.YOMITAN_LOOKUP) {
      bucket.lookups += 1;
    } else {
      bucket.cards += 1;
    }
  }

  return { bucketMs, startMs, endMs, hasKnownWords, buckets };
}
