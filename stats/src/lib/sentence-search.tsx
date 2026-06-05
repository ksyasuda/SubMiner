import type { ReactNode } from 'react';
import type { SentenceSearchResult } from '../types/stats';

export interface SentenceMatchRange {
  start: number;
  end: number;
}

export interface SentenceSearchMineAvailability {
  canMineSentence: boolean;
  canMineWordAudio: boolean;
  exactMatch: boolean;
  unavailableReason: string | null;
}

function normalizedSearchWord(query: string): string {
  return query.trim();
}

export function findExactSentenceMatches(text: string, query: string): SentenceMatchRange[] {
  const needle = normalizedSearchWord(query);
  if (!needle) return [];

  const ranges: SentenceMatchRange[] = [];
  const haystack = text.toLocaleLowerCase();
  const normalizedNeedle = needle.toLocaleLowerCase();
  let searchFrom = 0;

  while (searchFrom < haystack.length) {
    const index = haystack.indexOf(normalizedNeedle, searchFrom);
    if (index < 0) break;
    ranges.push({ start: index, end: index + needle.length });
    searchFrom = index + needle.length;
  }

  return ranges;
}

export function getSentenceSearchMineAvailability(
  result: Pick<SentenceSearchResult, 'sourcePath' | 'segmentStartMs' | 'segmentEndMs' | 'text'>,
  query: string,
): SentenceSearchMineAvailability {
  const exactMatch = findExactSentenceMatches(result.text, query).length > 0;
  const unavailableReason = !result.sourcePath
    ? 'This source has no local file path.'
    : result.segmentStartMs == null || result.segmentEndMs == null
      ? 'This line is missing segment timing.'
      : null;

  return {
    canMineSentence: unavailableReason === null,
    canMineWordAudio: unavailableReason === null && exactMatch,
    exactMatch,
    unavailableReason,
  };
}

export function renderSentenceWithMatches(text: string, query: string): ReactNode {
  const ranges = findExactSentenceMatches(text, query);
  if (ranges.length === 0) return text;

  const parts: ReactNode[] = [];
  let cursor = 0;
  ranges.forEach((range, index) => {
    if (range.start > cursor) {
      parts.push(text.slice(cursor, range.start));
    }
    parts.push(
      <mark
        key={`${range.start}-${index}`}
        className="rounded bg-ctp-yellow/15 px-0.5 text-ctp-yellow underline decoration-ctp-yellow/60 underline-offset-2"
      >
        {text.slice(range.start, range.end)}
      </mark>,
    );
    cursor = range.end;
  });

  if (cursor < text.length) {
    parts.push(text.slice(cursor));
  }

  return parts;
}
