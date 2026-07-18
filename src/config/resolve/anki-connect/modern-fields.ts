import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asString } from '../shared';
import { applyModernValue } from './modern-value';

export function applyModernFieldsResolution(
  context: ResolveContext,
  fields: Record<string, unknown>,
): void {
  for (const key of ['word', 'audio', 'image', 'sentence', 'miscInfo', 'translation'] as const) {
    applyModernValue(
      context,
      fields,
      key,
      `ankiConnect.fields.${key}`,
      asString,
      DEFAULT_CONFIG.ankiConnect.fields[key],
      (value) => {
        context.resolved.ankiConnect.fields[key] = value;
      },
      'Expected string.',
    );
  }
}
