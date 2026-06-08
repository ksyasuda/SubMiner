import type { SentenceSearchResult } from '../types/stats';

export type StatsMineMode = 'word' | 'sentence' | 'audio';

export interface StatsMineCardParams {
  sourcePath: string;
  startMs: number;
  endMs: number;
  sentence: string;
  word: string;
  secondaryText?: string | null;
  videoTitle: string;
  mode: StatsMineMode;
}

export interface StatsMineCardResponse {
  noteId?: number;
  error?: string;
  errors?: string[];
}

export function getStatsMineCardUnavailableReason(
  result: Pick<SentenceSearchResult, 'sourcePath' | 'segmentStartMs' | 'segmentEndMs'>,
): string | null {
  if (!result.sourcePath) {
    return 'This source has no local file path.';
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
