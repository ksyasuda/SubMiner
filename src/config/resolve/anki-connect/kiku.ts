import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';

export function applyAnkiKikuResolution(context: ResolveContext): void {
  if (
    context.resolved.ankiConnect.isKiku.fieldGrouping !== 'auto' &&
    context.resolved.ankiConnect.isKiku.fieldGrouping !== 'manual' &&
    context.resolved.ankiConnect.isKiku.fieldGrouping !== 'disabled'
  ) {
    context.warn(
      'ankiConnect.isKiku.fieldGrouping',
      context.resolved.ankiConnect.isKiku.fieldGrouping,
      DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping,
      'Expected auto, manual, or disabled.',
    );
    context.resolved.ankiConnect.isKiku.fieldGrouping =
      DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping;
  }
}
