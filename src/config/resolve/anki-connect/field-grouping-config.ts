import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asBoolean, isObject } from '../shared';
import { applyModernValue } from './modern-value';

type FieldGroupingConfigKey = 'isKiku' | 'isSenren';

export function applyFieldGroupingConfigResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
  key: FieldGroupingConfigKey,
): void {
  const source = ankiConnect[key];
  if (!isObject(source)) {
    if (source !== undefined) {
      context.warn(
        `ankiConnect.${key}`,
        source,
        DEFAULT_CONFIG.ankiConnect[key],
        'Expected object.',
      );
    }
    return;
  }

  for (const booleanKey of ['enabled', 'deleteDuplicateInAuto'] as const) {
    applyModernValue(
      context,
      source,
      booleanKey,
      `ankiConnect.${key}.${booleanKey}`,
      asBoolean,
      DEFAULT_CONFIG.ankiConnect[key][booleanKey],
      (value) => {
        context.resolved.ankiConnect[key][booleanKey] = value;
      },
      'Expected boolean.',
    );
  }

  applyModernValue(
    context,
    source,
    'fieldGrouping',
    `ankiConnect.${key}.fieldGrouping`,
    (value) => (value === 'auto' || value === 'manual' || value === 'disabled' ? value : undefined),
    DEFAULT_CONFIG.ankiConnect[key].fieldGrouping,
    (value) => {
      context.resolved.ankiConnect[key].fieldGrouping = value;
    },
    'Expected auto, manual, or disabled.',
  );
}
