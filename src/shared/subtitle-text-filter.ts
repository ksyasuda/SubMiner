import type { SubtitleCue, SubtitleFilterConfig, SubtitleFilterPresetId } from '../types/subtitle';

/**
 * Subtitle text filters (`subtitleFilter` config): built-in presets plus user regex rules,
 * applied to primary subtitle text as it enters the app (live mpv `sub-text` and parsed
 * cues), so the overlay, sidebar, tokenizer, mining and stats all see the same text.
 * Shared with the settings window, which validates custom patterns as they are typed.
 */

interface SubtitleFilterPreset {
  label: string;
  description: string;
  pattern: string;
}

/** Every rule compiles with these flags: global, per-line `^`/`$`, and Unicode escapes. */
export const SUBTITLE_FILTER_REGEX_FLAGS = 'gmu';

// One pair of parentheses without nesting or line breaks.
const PARENTHESIZED = '[（(][^（）()\\n]*[）)]';
// Horizontal whitespace, including the full-width space.
const LINE_SPACE = '[^\\S\\n]*';

export const SUBTITLE_FILTER_PRESETS: Record<SubtitleFilterPresetId, SubtitleFilterPreset> = {
  parentheses: {
    label: 'Text in Parentheses',
    description:
      'Remove text in （） and (), such as speaker names, readings and notes. A line wrapped entirely in parentheses, usually inner thoughts, is kept.',
    // Matches unless the parentheses both start and end the line.
    pattern: `(?<!^${LINE_SPACE})${PARENTHESIZED}|${PARENTHESIZED}(?!${LINE_SPACE}$)`,
  },
  brackets: {
    label: 'Text in Square Brackets',
    description: 'Remove text in ［］, [], 【】 and 〔〕, such as sound effects and music cues.',
    pattern: '[［\\[【〔][^［］\\[\\]【】〔〕\\n]*[］\\]】〕]',
  },
  speakerPrefix: {
    label: 'Speaker Names Before a Colon',
    description:
      'Remove a short name followed by a colon at the start of a line, as in 山田：元気？',
    pattern: '^[^\\s\\d:：（）()「」『』]{1,12}[：:]',
  },
  musicSymbols: {
    label: 'Music Notes',
    description: 'Remove ♪, ♫ and ♬. A line left with only symbols such as ～ is removed.',
    pattern: '[♪♫♬]+',
  },
};

export const SUBTITLE_FILTER_PRESET_IDS = Object.keys(
  SUBTITLE_FILTER_PRESETS,
) as SubtitleFilterPresetId[];

interface CompiledSubtitleFilterRule {
  regex: RegExp;
  replacement: string;
}

/** Returns the compile error for a custom rule pattern, or null when it is valid. */
export function getSubtitleFilterPatternError(pattern: string): string | null {
  if (!pattern) return 'Pattern is empty.';
  try {
    new RegExp(pattern, SUBTITLE_FILTER_REGEX_FLAGS);
    return null;
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
}

// Config objects are replaced, never mutated, on reload, so identity keys the cache.
const compiledRulesCache = new WeakMap<SubtitleFilterConfig, CompiledSubtitleFilterRule[]>();

function compileSubtitleFilterRules(config: SubtitleFilterConfig): CompiledSubtitleFilterRule[] {
  const cached = compiledRulesCache.get(config);
  if (cached) return cached;

  const rules: CompiledSubtitleFilterRule[] = [];
  for (const id of SUBTITLE_FILTER_PRESET_IDS) {
    if (config.presets?.[id] === true) {
      const regex = new RegExp(SUBTITLE_FILTER_PRESETS[id].pattern, SUBTITLE_FILTER_REGEX_FLAGS);
      rules.push({ regex, replacement: '' });
    }
  }
  for (const rule of config.customRules ?? []) {
    if (rule.enabled === false || !rule.pattern) continue;
    // Invalid patterns are reported when the config resolves; here they are skipped.
    if (getSubtitleFilterPatternError(rule.pattern) !== null) continue;
    rules.push({
      regex: new RegExp(rule.pattern, SUBTITLE_FILTER_REGEX_FLAGS),
      replacement: rule.replacement ?? '',
    });
  }
  compiledRulesCache.set(config, rules);
  return rules;
}

export function hasActiveSubtitleFilters(config: SubtitleFilterConfig | undefined): boolean {
  return config !== undefined && compileSubtitleFilterRules(config).length > 0;
}

const HAS_CONTENT = /[\p{L}\p{N}]/u;

/**
 * Tidies the lines a filter changed: collapses the double space a removed word leaves,
 * trims the line, and drops it when only punctuation or symbols remain. Lines the filter
 * did not touch are kept as authored, so a line of just "……" survives.
 */
function tidyFilteredLines(original: string, filtered: string): string {
  const before = original.split('\n');
  const after = filtered.split('\n');
  const sameShape = before.length === after.length;
  const lines: string[] = [];
  after.forEach((line, index) => {
    if (sameShape && line === before[index]) {
      lines.push(line);
      return;
    }
    const tidied = line.replace(/ {2,}/g, ' ').trim();
    if (HAS_CONTENT.test(tidied)) lines.push(tidied);
  });
  return lines.join('\n').replace(/^\n+|\n+$/g, '');
}

/**
 * Applies the enabled presets, then custom rules in order. Returns '' for an emptied line.
 * Lines are tidied after each rule, so a later `^` anchor is not blocked by the space an
 * earlier rule left at the start of a line.
 */
export function filterSubtitleText(text: string, config: SubtitleFilterConfig | undefined): string {
  if (!text || !config) return text;
  let filtered = text;
  for (const rule of compileSubtitleFilterRules(config)) {
    const next = filtered.replace(rule.regex, rule.replacement);
    if (next !== filtered) filtered = tidyFilteredLines(filtered, next);
  }
  return filtered;
}

/**
 * Filters freshly parsed cues in place and drops cues the filter emptied. Cue objects are
 * mutated rather than copied because the parser attaches non-enumerable metadata.
 */
export function filterSubtitleCues(
  cues: SubtitleCue[],
  config: SubtitleFilterConfig | undefined,
): SubtitleCue[] {
  if (!hasActiveSubtitleFilters(config)) return cues;
  return cues.filter((cue) => {
    cue.text = filterSubtitleText(cue.text, config);
    return cue.text.length > 0;
  });
}
