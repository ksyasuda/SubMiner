import type {
  SentenceSearchResult,
  StatsMineCardParams,
  StatsMineCardResponse,
  StatsMineMode,
} from '../types/stats';

export type { StatsMineCardParams, StatsMineCardResponse, StatsMineMode } from '../types/stats';

// Cards cut their audio and image from the video file. Only streamed videos have no
// local path, so a missing path means the line came from a stream.
const STREAMED_LINE_REASON =
  "Can't mine: from a stream, so there's no video file to cut clips from.";

/** Explains how a word's lines affect mining, or null when every line can be mined. */
export function describeStreamedLines(lineCount: number, streamLineCount: number): string | null {
  if (streamLineCount === 0) return null;
  if (streamLineCount >= lineCount) {
    return lineCount === 1
      ? "This word's only line came from a stream, so it can't be mined yet."
      : `All ${lineCount.toLocaleString()} lines with this word came from streams, so it can't be mined yet.`;
  }
  return `${streamLineCount.toLocaleString()} of ${lineCount.toLocaleString()} lines with this word came from streams and can't be mined.`;
}

export function getStatsMineCardUnavailableReason(
  result: Pick<SentenceSearchResult, 'sourcePath' | 'segmentStartMs' | 'segmentEndMs'>,
): string | null {
  if (!result.sourcePath) {
    return STREAMED_LINE_REASON;
  }
  if (result.segmentStartMs == null || result.segmentEndMs == null) {
    return 'This line is missing segment timing.';
  }
  if (
    !Number.isFinite(result.segmentStartMs) ||
    !Number.isFinite(result.segmentEndMs) ||
    result.segmentEndMs <= result.segmentStartMs
  ) {
    return 'This line has invalid segment timing.';
  }
  return null;
}

export function buildStatsMineCardParams(
  result: Pick<
    SentenceSearchResult,
    'sourcePath' | 'segmentStartMs' | 'segmentEndMs' | 'text' | 'secondaryText' | 'videoTitle'
  >,
  word: string,
  mode: StatsMineMode,
): StatsMineCardParams | null {
  if (getStatsMineCardUnavailableReason(result)) {
    return null;
  }

  return {
    sourcePath: result.sourcePath!,
    startMs: result.segmentStartMs!,
    endMs: result.segmentEndMs!,
    sentence: result.text,
    word,
    secondaryText: result.secondaryText,
    videoTitle: result.videoTitle,
    mode,
  };
}

export function getStatsMineCardError(response: StatsMineCardResponse): string | null {
  if (response.error) return response.error;
  return response.errors?.[0] ?? null;
}
