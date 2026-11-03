import type { SubtitleFilterRule } from '../../types/subtitle';
import {
  SUBTITLE_FILTER_PRESET_IDS,
  getSubtitleFilterPatternError,
} from '../../shared/subtitle-text-filter';
import { ResolveContext } from './context';
import { asBoolean, asString, isObject } from './shared';

function resolveCustomRule(
  raw: unknown,
  index: number,
  warn: ResolveContext['warn'],
): SubtitleFilterRule | null {
  const path = `subtitleFilter.customRules[${index}]`;
  const pattern = isObject(raw) ? asString(raw.pattern) : undefined;
  if (!isObject(raw) || pattern === undefined) {
    warn(path, raw, null, 'Expected an object with a string pattern.');
    return null;
  }

  // Empty and invalid patterns stay in the config so the settings window can show and
  // fix them; the filter skips both at runtime.
  const patternError = pattern ? getSubtitleFilterPatternError(pattern) : null;
  if (patternError !== null) {
    warn(`${path}.pattern`, pattern, null, `Invalid regular expression: ${patternError}`);
  }

  return {
    name: asString(raw.name) ?? '',
    pattern,
    replacement: asString(raw.replacement) ?? '',
    enabled: asBoolean(raw.enabled) ?? true,
  };
}

export function applySubtitleFilterConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;
  if (src.subtitleFilter === undefined) return;
  if (!isObject(src.subtitleFilter)) {
    warn('subtitleFilter', src.subtitleFilter, resolved.subtitleFilter, 'Expected object.');
    return;
  }

  const { presets, customRules } = src.subtitleFilter;
  if (isObject(presets)) {
    for (const id of SUBTITLE_FILTER_PRESET_IDS) {
      const enabled = asBoolean(presets[id]);
      if (enabled !== undefined) {
        resolved.subtitleFilter.presets[id] = enabled;
      } else if (presets[id] !== undefined) {
        warn(
          `subtitleFilter.presets.${id}`,
          presets[id],
          resolved.subtitleFilter.presets[id],
          'Expected boolean.',
        );
      }
    }
  } else if (presets !== undefined) {
    warn('subtitleFilter.presets', presets, resolved.subtitleFilter.presets, 'Expected object.');
  }

  if (Array.isArray(customRules)) {
    resolved.subtitleFilter.customRules = customRules
      .map((rule, index) => resolveCustomRule(rule, index, warn))
      .filter((rule): rule is SubtitleFilterRule => rule !== null);
  } else if (customRules !== undefined) {
    warn(
      'subtitleFilter.customRules',
      customRules,
      resolved.subtitleFilter.customRules,
      'Expected an array of rules.',
    );
  }
}
