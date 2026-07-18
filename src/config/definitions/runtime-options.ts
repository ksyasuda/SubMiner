import { ResolvedConfig } from '../../types/config';
import { RuntimeOptionRegistryEntry } from './shared';

export function buildRuntimeOptionRegistry(
  defaultConfig: ResolvedConfig,
): RuntimeOptionRegistryEntry[] {
  return [
    {
      id: 'anki.autoUpdateNewCards',
      path: 'ankiConnect.behavior.autoUpdateNewCards',
      label: 'Auto Update New Cards',
      scope: 'ankiConnect',
      valueType: 'boolean',
      allowedValues: [true, false],
      defaultValue: defaultConfig.ankiConnect.behavior.autoUpdateNewCards,
      requiresRestart: false,
      formatValueForOsd: (value) => (value === true ? 'On' : 'Off'),
      toAnkiPatch: (value) => ({
        behavior: { autoUpdateNewCards: value === true },
      }),
    },
    {
      id: 'subtitle.annotation.knownWords.highlightEnabled',
      path: 'ankiConnect.knownWords.highlightEnabled',
      label: 'Known Word Annotation',
      scope: 'subtitle',
      valueType: 'boolean',
      allowedValues: [true, false],
      defaultValue: defaultConfig.ankiConnect.knownWords.highlightEnabled,
      requiresRestart: false,
      formatValueForOsd: (value) => (value === true ? 'On' : 'Off'),
      toAnkiPatch: (value) => ({
        knownWords: {
          highlightEnabled: value === true,
        },
      }),
    },
    {
      id: 'subtitle.annotation.knownWords.maturityEnabled',
      path: 'ankiConnect.knownWords.maturityEnabled',
      label: 'Known Word Maturity Colors',
      scope: 'subtitle',
      valueType: 'boolean',
      allowedValues: [true, false],
      defaultValue: defaultConfig.ankiConnect.knownWords.maturityEnabled,
      requiresRestart: false,
      formatValueForOsd: (value) => (value === true ? 'On' : 'Off'),
      toAnkiPatch: (value) => ({
        knownWords: {
          maturityEnabled: value === true,
        },
      }),
    },
    {
      id: 'subtitle.annotation.nPlusOne',
      path: 'ankiConnect.nPlusOne.enabled',
      label: 'N+1 Annotation',
      scope: 'subtitle',
      valueType: 'boolean',
      allowedValues: [true, false],
      defaultValue: defaultConfig.ankiConnect.nPlusOne.enabled,
      requiresRestart: false,
      formatValueForOsd: (value) => (value === true ? 'On' : 'Off'),
      toAnkiPatch: (value) => ({
        nPlusOne: {
          enabled: value === true,
        },
      }),
    },
    {
      id: 'subtitle.annotation.jlpt',
      path: 'subtitleStyle.enableJlpt',
      label: 'JLPT Annotation',
      scope: 'subtitle',
      valueType: 'boolean',
      allowedValues: [true, false],
      defaultValue: defaultConfig.subtitleStyle.enableJlpt,
      requiresRestart: false,
      formatValueForOsd: (value) => (value === true ? 'On' : 'Off'),
      toAnkiPatch: () => ({}),
    },
    {
      id: 'subtitle.annotation.frequency',
      path: 'subtitleStyle.frequencyDictionary.enabled',
      label: 'Frequency Annotation',
      scope: 'subtitle',
      valueType: 'boolean',
      allowedValues: [true, false],
      defaultValue: defaultConfig.subtitleStyle.frequencyDictionary.enabled,
      requiresRestart: false,
      formatValueForOsd: (value) => (value === true ? 'On' : 'Off'),
      toAnkiPatch: () => ({}),
    },
    {
      id: 'anki.nPlusOneMatchMode',
      path: 'ankiConnect.knownWords.matchMode',
      label: 'Known Word Match Mode',
      scope: 'ankiConnect',
      valueType: 'enum',
      allowedValues: ['headword', 'surface'],
      defaultValue: defaultConfig.ankiConnect.knownWords.matchMode,
      requiresRestart: false,
      formatValueForOsd: (value) => String(value),
      toAnkiPatch: (value) => ({
        knownWords: {
          matchMode: value === 'headword' || value === 'surface' ? value : 'headword',
        },
      }),
    },
    {
      id: 'anki.kikuFieldGrouping',
      path: 'ankiConnect.isKiku.fieldGrouping',
      label: 'Kiku Field Grouping',
      scope: 'ankiConnect',
      valueType: 'enum',
      allowedValues: ['auto', 'manual', 'disabled'],
      defaultValue: 'disabled',
      requiresRestart: false,
      formatValueForOsd: (value) => String(value),
      toAnkiPatch: (value) => ({
        isKiku: {
          fieldGrouping:
            value === 'auto' || value === 'manual' || value === 'disabled' ? value : 'disabled',
        },
      }),
    },
  ];
}
