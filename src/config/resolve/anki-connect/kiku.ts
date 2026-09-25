import type { ResolveContext } from '../context';
import { applyFieldGroupingConfigResolution } from './field-grouping-config';

export function applyAnkiKikuResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  applyFieldGroupingConfigResolution(context, ankiConnect, 'isKiku');
}
