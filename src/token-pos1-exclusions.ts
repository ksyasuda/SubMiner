import type { ResolvedTokenPos1ExclusionConfig, TokenPos1ExclusionConfig } from './types';

export const DEFAULT_ANNOTATION_POS1_EXCLUSION_DEFAULTS = Object.freeze([
  '助詞',
  '助動詞',
  '記号',
  '補助記号',
  '連体詞',
  '感動詞',
  '接続詞',
  '接頭詞',
]) as readonly string[];

export const DEFAULT_ANNOTATION_POS1_EXCLUSION_CONFIG: ResolvedTokenPos1ExclusionConfig = {
  defaults: [...DEFAULT_ANNOTATION_POS1_EXCLUSION_DEFAULTS],
  add: [],
  remove: [],
};

function normalizePosTag(value: string): string {
  return value.trim();
}

export function normalizePos1ExclusionList(values: readonly string[]): string[] {
  const deduped = new Set<string>();
  for (const value of values) {
    const normalized = normalizePosTag(value);
    if (!normalized) {
      continue;
    }
    deduped.add(normalized);
  }
  return [...deduped];
}

export function resolveAnnotationPos1ExclusionSet(
  config: TokenPos1ExclusionConfig | ResolvedTokenPos1ExclusionConfig,
): ReadonlySet<string> {
  const defaults = normalizePos1ExclusionList(config.defaults ?? []);
  const added = normalizePos1ExclusionList(config.add ?? []);
  const removed = new Set(normalizePos1ExclusionList(config.remove ?? []));
  const resolved = new Set<string>();
  for (const value of defaults) {
    resolved.add(value);
  }
  for (const value of added) {
    resolved.add(value);
  }
  for (const value of removed) {
    resolved.delete(value);
  }
  return resolved;
}
