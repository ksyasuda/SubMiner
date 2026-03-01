import type { ResolvedTokenPos2ExclusionConfig, TokenPos2ExclusionConfig } from './types';
import { normalizePos1ExclusionList } from './token-pos1-exclusions';

export const DEFAULT_ANNOTATION_POS2_EXCLUSION_DEFAULTS = Object.freeze(['非自立']) as readonly string[];

export const DEFAULT_ANNOTATION_POS2_EXCLUSION_CONFIG: ResolvedTokenPos2ExclusionConfig = {
  defaults: [...DEFAULT_ANNOTATION_POS2_EXCLUSION_DEFAULTS],
  add: [],
  remove: [],
};

export function resolveAnnotationPos2ExclusionSet(
  config: TokenPos2ExclusionConfig | ResolvedTokenPos2ExclusionConfig,
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
