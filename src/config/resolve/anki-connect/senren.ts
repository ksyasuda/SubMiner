import type { ResolveContext } from '../context';
import { applyFieldGroupingConfigResolution } from './field-grouping-config';

export function applyAnkiSenrenResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  applyFieldGroupingConfigResolution(context, ankiConnect, 'isSenren');

  // Kiku and Senren field grouping write incompatible markup into the same note
  // fields, so only one may be active; Kiku wins to preserve pre-existing setups.
  if (
    context.resolved.ankiConnect.isSenren.enabled === true &&
    context.resolved.ankiConnect.isKiku.enabled === true
  ) {
    context.warn(
      'ankiConnect.isSenren.enabled',
      true,
      false,
      'Kiku and Senren are mutually exclusive; disable isKiku.enabled to use Senren field grouping.',
    );
    context.resolved.ankiConnect.isSenren.enabled = false;
  }
}
