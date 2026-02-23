import { ResolvedConfig } from '../../types';
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
      id: 'anki.nPlusOneMatchMode',
      path: 'ankiConnect.nPlusOne.matchMode',
      label: 'N+1 Match Mode',
      scope: 'ankiConnect',
      valueType: 'enum',
      allowedValues: ['headword', 'surface'],
      defaultValue: defaultConfig.ankiConnect.nPlusOne.matchMode,
      requiresRestart: false,
      formatValueForOsd: (value) => String(value),
      toAnkiPatch: (value) => ({
        nPlusOne: {
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
