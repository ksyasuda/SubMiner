import { ResolvedConfig } from '../../types';
import { ConfigOptionRegistryEntry } from './shared';

export function buildSubtitleConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
): ConfigOptionRegistryEntry[] {
  return [
    {
      path: 'subtitleStyle.enableJlpt',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleStyle.enableJlpt,
      description:
        'Enable JLPT vocabulary level underlines. ' +
        'When disabled, JLPT tagging lookup and underlines are skipped.',
    },
    {
      path: 'subtitleStyle.preserveLineBreaks',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleStyle.preserveLineBreaks,
      description:
        'Preserve line breaks in visible overlay subtitle rendering. ' +
        'When false, line breaks are flattened to spaces for a single-line flow.',
    },
    {
      path: 'subtitleStyle.autoPauseVideoOnHover',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleStyle.autoPauseVideoOnHover,
      description:
        'Automatically pause mpv playback while hovering subtitle text, then resume on leave.',
    },
    {
      path: 'subtitleStyle.autoPauseVideoOnYomitanPopup',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleStyle.autoPauseVideoOnYomitanPopup,
      description:
        'Automatically pause mpv playback while Yomitan popup is open, then resume when popup closes.',
    },
    {
      path: 'subtitleStyle.hoverTokenColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleStyle.hoverTokenColor,
      description: 'Hex color used for hovered subtitle token highlight in mpv.',
    },
    {
      path: 'subtitleStyle.hoverTokenBackgroundColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleStyle.hoverTokenBackgroundColor,
      description: 'CSS color used for hovered subtitle token background highlight in mpv.',
    },
    {
      path: 'subtitleStyle.frequencyDictionary.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.enabled,
      description: 'Enable frequency-dictionary-based highlighting based on token rank.',
    },
    {
      path: 'subtitleStyle.frequencyDictionary.sourcePath',
      kind: 'string',
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.sourcePath,
      description:
        'Optional absolute path to a frequency dictionary directory.' +
        ' If empty, built-in discovery search paths are used.',
    },
    {
      path: 'subtitleStyle.frequencyDictionary.topX',
      kind: 'number',
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.topX,
      description: 'Only color tokens with frequency rank <= topX (default: 1000).',
    },
    {
      path: 'subtitleStyle.frequencyDictionary.mode',
      kind: 'enum',
      enumValues: ['single', 'banded'],
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.mode,
      description:
        'single: use one color for all matching tokens. banded: use color ramp by frequency band.',
    },
    {
      path: 'subtitleStyle.frequencyDictionary.matchMode',
      kind: 'enum',
      enumValues: ['headword', 'surface'],
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.matchMode,
      description:
        'headword: frequency lookup uses dictionary form. surface: lookup uses subtitle-visible token text.',
    },
    {
      path: 'subtitleStyle.frequencyDictionary.singleColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.singleColor,
      description: 'Color used when frequencyDictionary.mode is `single`.',
    },
    {
      path: 'subtitleStyle.frequencyDictionary.bandedColors',
      kind: 'array',
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.bandedColors,
      description:
        'Five colors used for rank bands when mode is `banded` (from most common to least within topX).',
    },
  ];
}
