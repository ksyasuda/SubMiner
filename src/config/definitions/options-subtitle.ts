import { ResolvedConfig } from '../../types/config';
import { ConfigOptionRegistryEntry } from './shared';

export function buildSubtitleConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
): ConfigOptionRegistryEntry[] {
  return [
    {
      path: 'subtitleStyle.primaryDefaultMode',
      kind: 'enum',
      enumValues: ['hidden', 'visible', 'hover'],
      defaultValue: defaultConfig.subtitleStyle.primaryDefaultMode,
      description:
        'Default primary subtitle bar visibility mode. hidden hides it, visible shows it, hover reveals it on hover.',
    },
    {
      path: 'subtitleStyle.css',
      kind: 'object',
      defaultValue: defaultConfig.subtitleStyle.css,
      description:
        'CSS declaration object applied to primary subtitles after normal subtitle style defaults.',
    },
    {
      path: 'subtitleStyle.secondary.css',
      kind: 'object',
      defaultValue: defaultConfig.subtitleStyle.secondary.css,
      description:
        'CSS declaration object applied to secondary subtitles after normal subtitle style defaults.',
    },
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
      path: 'subtitleStyle.nameMatchEnabled',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleStyle.nameMatchEnabled,
      description:
        'Enable character dictionary sync and subtitle token coloring for character-name matches.',
    },
    {
      path: 'subtitleStyle.nameMatchImagesEnabled',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleStyle.nameMatchImagesEnabled,
      description:
        'Show small character portraits beside subtitle tokens matched from the SubMiner character dictionary.',
    },
    {
      path: 'subtitleStyle.nameMatchColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleStyle.nameMatchColor,
      description:
        'Hex color used when a subtitle token matches an entry from the SubMiner character dictionary.',
    },
    {
      path: 'subtitleStyle.knownWordColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleStyle.knownWordColor,
      description: 'Color used for known-word subtitle highlights.',
    },
    {
      path: 'subtitleStyle.nPlusOneColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleStyle.nPlusOneColor,
      description: 'Color used for the single N+1 target token subtitle highlight.',
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
    {
      path: 'subtitleSidebar.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleSidebar.enabled,
      description: 'Enable the subtitle sidebar feature for parsed subtitle sources.',
    },
    {
      path: 'subtitleSidebar.autoOpen',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleSidebar.autoOpen,
      description: 'Automatically open the subtitle sidebar once during overlay startup.',
    },
    {
      path: 'subtitleSidebar.layout',
      kind: 'enum',
      enumValues: ['overlay', 'embedded'],
      defaultValue: defaultConfig.subtitleSidebar.layout,
      description: 'Render the subtitle sidebar as a floating overlay or reserve space inside mpv.',
    },
    {
      path: 'subtitleSidebar.toggleKey',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.toggleKey,
      description: 'KeyboardEvent.code used to toggle the subtitle sidebar open and closed.',
    },
    {
      path: 'subtitleSidebar.pauseVideoOnHover',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleSidebar.pauseVideoOnHover,
      description: 'Pause mpv while hovering the subtitle sidebar, then resume on leave.',
    },
    {
      path: 'subtitleSidebar.autoScroll',
      kind: 'boolean',
      defaultValue: defaultConfig.subtitleSidebar.autoScroll,
      description: 'Auto-scroll the active subtitle cue into view while playback advances.',
    },
    {
      path: 'subtitleSidebar.css',
      kind: 'object',
      defaultValue: defaultConfig.subtitleSidebar.css,
      description:
        'CSS declaration object applied to the subtitle sidebar. Includes color, background-color, and all font properties.',
    },
    {
      path: 'subtitleSidebar.maxWidth',
      kind: 'number',
      defaultValue: defaultConfig.subtitleSidebar.maxWidth,
      description: 'Maximum sidebar width in CSS pixels.',
    },
    {
      path: 'subtitleSidebar.opacity',
      kind: 'number',
      defaultValue: defaultConfig.subtitleSidebar.opacity,
      description: 'Base opacity applied to the sidebar shell.',
    },
    {
      path: 'subtitleSidebar.backgroundColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.backgroundColor,
      description: 'Background color for the subtitle sidebar shell.',
    },
    {
      path: 'subtitleSidebar.textColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.textColor,
      description: 'Default cue text color in the subtitle sidebar.',
    },
    {
      path: 'subtitleSidebar.fontFamily',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.fontFamily,
      description: 'Font family used for subtitle sidebar cue text.',
    },
    {
      path: 'subtitleSidebar.fontSize',
      kind: 'number',
      defaultValue: defaultConfig.subtitleSidebar.fontSize,
      description: 'Base font size for subtitle sidebar cue text in CSS pixels.',
    },
    {
      path: 'subtitleSidebar.timestampColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.timestampColor,
      description: 'Timestamp color in the subtitle sidebar.',
    },
    {
      path: 'subtitleSidebar.activeLineColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.activeLineColor,
      description: 'Text color for the active subtitle cue.',
    },
    {
      path: 'subtitleSidebar.activeLineBackgroundColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.activeLineBackgroundColor,
      description: 'Background color for the active subtitle cue.',
    },
    {
      path: 'subtitleSidebar.hoverLineBackgroundColor',
      kind: 'string',
      defaultValue: defaultConfig.subtitleSidebar.hoverLineBackgroundColor,
      description: 'Background color for hovered subtitle cues.',
    },
  ];
}
