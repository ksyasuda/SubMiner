// On-disk shape of the known-word cache, plus the only parser for it.
//
// Two processes read this file: the cache manager (which rebuilds its indexes
// from it) and the stats server (which counts known words). They used to carry
// separate hand-written parsers, so bumping the format to v4 for maturity tiers
// left the stats server silently reporting zero known words. Everything that
// touches the format now goes through here, and the version dispatch below ends
// in assertNever so adding a V5 to the union fails the build at every consumer
// instead of degrading to an empty result at runtime.

import type { KnownWordMaturityTier } from '../types/subtitle';
import type { KnownWordEntry } from './known-word-entries';

export interface KnownWordCacheStateV1 {
  readonly version: 1;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly words: string[];
}

export interface KnownWordCacheStateV2 {
  readonly version: 2;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly words: string[];
  readonly notes: Record<string, string[]>;
}

export interface KnownWordCacheStateV3 {
  readonly version: 3;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly notes: Record<string, KnownWordEntry[]>;
}

export interface KnownWordCacheStateV4 {
  readonly version: 4;
  readonly refreshedAtMs: number;
  readonly scope: string;
  readonly notes: Record<string, KnownWordEntry[]>;
  readonly tiers: Record<string, KnownWordMaturityTier>;
}

export type KnownWordCacheState =
  | KnownWordCacheStateV1
  | KnownWordCacheStateV2
  | KnownWordCacheStateV3
  | KnownWordCacheStateV4;

// Version written by persistKnownWordCacheState. Readers accept every version
// in the union above; only the writer pins one.
export type CurrentKnownWordCacheState = KnownWordCacheStateV4;

// Exported so every consumer that switches on `version` can close its dispatch
// the same way: a new member of the union becomes a type error at each call
// site rather than a case that silently falls through.
export function assertNever(value: never): never {
  throw new Error(`Unhandled known-word cache state: ${JSON.stringify(value)}`);
}

function isEntryRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isKnownWordEntry(entry: unknown): boolean {
  if (!isEntryRecord(entry)) return false;
  const candidate = entry as Partial<KnownWordEntry>;
  return (
    typeof candidate.word === 'string' &&
    (candidate.reading === null || typeof candidate.reading === 'string')
  );
}

// Returns the narrowed state, or null when the payload is not a cache state we
// recognize. Per-entry values that are merely unusable (an unknown maturity
// tier, a non-numeric note id) are dropped by callers at load time rather than
// rejecting the whole file.
export function parseKnownWordCacheState(value: unknown): KnownWordCacheState | null {
  if (!isEntryRecord(value)) return null;
  const candidate = value;
  if (
    candidate.version !== 1 &&
    candidate.version !== 2 &&
    candidate.version !== 3 &&
    candidate.version !== 4
  ) {
    return null;
  }
  if (typeof candidate.refreshedAtMs !== 'number') return null;
  if (typeof candidate.scope !== 'string') return null;

  if (candidate.version === 1 || candidate.version === 2) {
    if (!Array.isArray(candidate.words)) return null;
    if (!candidate.words.every((entry: unknown) => typeof entry === 'string')) return null;
  }

  if (candidate.version === 4) {
    // Per-tier values are sanitized entry-by-entry at load time.
    if (!isEntryRecord(candidate.tiers)) return null;
  }

  if (candidate.version === 2 || candidate.version === 3 || candidate.version === 4) {
    if (!isEntryRecord(candidate.notes)) return null;
    const isValidNoteEntry =
      candidate.version === 2
        ? (entry: unknown): boolean => typeof entry === 'string'
        : isKnownWordEntry;
    if (
      !Object.values(candidate.notes).every(
        (noteEntries) => Array.isArray(noteEntries) && noteEntries.every(isValidNoteEntry),
      )
    ) {
      return null;
    }
  }

  return candidate as unknown as KnownWordCacheState;
}

// Every word the cache considers known, flattened across notes. Consumers that
// only need membership (the stats server) use this instead of walking the
// version-specific layout themselves.
export function knownWordsFromState(state: KnownWordCacheState): Set<string> {
  switch (state.version) {
    case 1:
    case 2:
      return new Set(state.words);
    case 3:
    case 4: {
      const words = new Set<string>();
      for (const entries of Object.values(state.notes)) {
        for (const entry of entries) {
          if (entry.word) words.add(entry.word);
        }
      }
      return words;
    }
    default:
      return assertNever(state);
  }
}
