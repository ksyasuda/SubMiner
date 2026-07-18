import type { AnkiConnectConfig } from '../types/anki';
import type { KnownWordMaturityTier } from '../types/subtitle';

export const DEFAULT_MATURE_INTERVAL_THRESHOLD_DAYS = 21;

// Ascending maturity; index order backs tier comparison.
const TIER_ORDER: readonly KnownWordMaturityTier[] = ['new', 'learning', 'young', 'mature'];

export interface KnownWordMaturityTierQueries {
  mature: string;
  young: string;
  learning: string;
}

export interface KnownWordMaturityTierSets {
  mature: ReadonlySet<number>;
  young: ReadonlySet<number>;
  learning: ReadonlySet<number>;
}

// Maturity tiers only affect how known-word highlights render, so both flags
// must be on before tier data is fetched or served.
export function getKnownWordMaturityEnabled(config: AnkiConnectConfig): boolean {
  return (
    config.knownWords?.highlightEnabled === true && config.knownWords?.maturityEnabled === true
  );
}

export function getMatureIntervalThresholdDays(config: AnkiConnectConfig): number {
  const threshold = config.knownWords?.matureThresholdDays;
  if (typeof threshold === 'number' && Number.isFinite(threshold) && threshold >= 1) {
    return Math.floor(threshold);
  }
  return DEFAULT_MATURE_INTERVAL_THRESHOLD_DAYS;
}

// Anki search props classify notes server-side: a note matches a tier query
// when ANY of its cards matches, which implements most-mature-card-wins for
// free once tiers are checked in mature > young > learning order.
export function buildKnownWordMaturityTierQueries(
  scopeQuery: string,
  thresholdDays: number,
): KnownWordMaturityTierQueries {
  const prefix = scopeQuery.trim().length > 0 ? `${scopeQuery.trim()} ` : '';
  return {
    mature: `${prefix}prop:ivl>=${thresholdDays}`,
    young: `${prefix}prop:ivl>=1 prop:ivl<${thresholdDays}`,
    learning: `${prefix}is:learn`,
  };
}

export async function fetchKnownWordMaturityTierSets(
  findNotes: (query: string, options?: { maxRetries?: number }) => Promise<unknown>,
  scopeQueries: string[],
  thresholdDays: number,
): Promise<{ mature: Set<number>; young: Set<number>; learning: Set<number> }> {
  const sets = {
    mature: new Set<number>(),
    young: new Set<number>(),
    learning: new Set<number>(),
  };
  for (const scopeQuery of scopeQueries) {
    const queries = buildKnownWordMaturityTierQueries(scopeQuery, thresholdDays);
    for (const tier of ['mature', 'young', 'learning'] as const) {
      const noteIds = (await findNotes(queries[tier], { maxRetries: 0 })) as number[];
      if (!Array.isArray(noteIds)) {
        continue;
      }
      for (const noteId of noteIds) {
        if (Number.isInteger(noteId) && noteId > 0) {
          sets[tier].add(noteId);
        }
      }
    }
  }
  return sets;
}

export function classifyKnownWordNoteTier(
  noteId: number,
  sets: KnownWordMaturityTierSets,
): KnownWordMaturityTier {
  if (sets.mature.has(noteId)) return 'mature';
  if (sets.young.has(noteId)) return 'young';
  if (sets.learning.has(noteId)) return 'learning';
  return 'new';
}

export function maxKnownWordMaturityTier(
  a: KnownWordMaturityTier | null | undefined,
  b: KnownWordMaturityTier | null | undefined,
): KnownWordMaturityTier | null {
  if (!a) return b ?? null;
  if (!b) return a;
  return TIER_ORDER.indexOf(a) >= TIER_ORDER.indexOf(b) ? a : b;
}

export function sanitizeKnownWordMaturityTier(value: unknown): KnownWordMaturityTier | null {
  return typeof value === 'string' && TIER_ORDER.includes(value as KnownWordMaturityTier)
    ? (value as KnownWordMaturityTier)
    : null;
}
