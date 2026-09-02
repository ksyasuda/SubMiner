import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';

export function applyAnkiSenrenResolution(context: ResolveContext): void {
  if (
    context.resolved.ankiConnect.isSenren.fieldGrouping !== 'auto' &&
    context.resolved.ankiConnect.isSenren.fieldGrouping !== 'manual' &&
    context.resolved.ankiConnect.isSenren.fieldGrouping !== 'disabled'
  ) {
    context.warn(
      'ankiConnect.isSenren.fieldGrouping',
      context.resolved.ankiConnect.isSenren.fieldGrouping,
      DEFAULT_CONFIG.ankiConnect.isSenren.fieldGrouping,
      'Expected auto, manual, or disabled.',
    );
    context.resolved.ankiConnect.isSenren.fieldGrouping =
      DEFAULT_CONFIG.ankiConnect.isSenren.fieldGrouping;
  }

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
