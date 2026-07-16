import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asString } from '../shared';
import { applyModernValue } from './modern-value';

export function applyModernMetadataResolution(
  context: ResolveContext,
  metadata: Record<string, unknown>,
): void {
  applyModernValue(
    context,
    metadata,
    'pattern',
    'ankiConnect.metadata.pattern',
    asString,
    DEFAULT_CONFIG.ankiConnect.metadata.pattern,
    (value) => {
      context.resolved.ankiConnect.metadata.pattern = value;
    },
    'Expected string.',
  );
}
