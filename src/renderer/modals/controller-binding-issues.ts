import { controllerReviewInputSourceKey } from '../../shared/controller-review';
import { DPAD_AXIS_INDEX, DPAD_BUTTON_INDEX } from '../handlers/gamepad-controller.js';
import type {
  ControllerDpadFallback,
  ResolvedControllerAxisBinding,
  ResolvedControllerConfig,
  ResolvedControllerDiscreteBinding,
} from '../../types';

type ControllerBindings = ResolvedControllerConfig['bindings'];
type ControllerBindingActionId = keyof ControllerBindings;

/** What the checker needs from an Alt+C row definition. */
export type ControllerBindingIssueDefinition = {
  id: ControllerBindingActionId;
  label: string;
  bindingType: 'discrete' | 'axis';
};

/** Row key of a stick binding's D-pad row in Alt+C. */
export function getControllerDpadRowKey(actionId: ControllerBindingActionId): string {
  return `${actionId}:dpad`;
}

const buttonKey = (buttonIndex: number) => `button:${buttonIndex}`;
const axisKey = (axisIndex: number, direction: 'negative' | 'positive') =>
  `axis:${axisIndex}:${direction}`;

/**
 * A combo counts as one input, apart from its hold and press alone: a plain action on the hold
 * button waits for release, and a combo keeps its press away from plain actions and stick rows.
 */
function discreteInputs(binding: ResolvedControllerDiscreteBinding): string[] {
  if (binding.kind === 'none') return [];
  const press = controllerReviewInputSourceKey(binding);
  return [binding.hold ? `${controllerReviewInputSourceKey(binding.hold)}+${press}` : press];
}

function stickInputs(binding: ResolvedControllerAxisBinding): string[] {
  if (binding.kind === 'none') return [];
  return [axisKey(binding.axisIndex, 'negative'), axisKey(binding.axisIndex, 'positive')];
}

/** The D-pad is read from buttons 12-15 and, on some raw pads, from a hat axis. */
function dpadInputs(fallback: ControllerDpadFallback): string[] {
  if (fallback === 'horizontal') {
    return [
      buttonKey(DPAD_BUTTON_INDEX.left),
      buttonKey(DPAD_BUTTON_INDEX.right),
      axisKey(DPAD_AXIS_INDEX.horizontal, 'negative'),
      axisKey(DPAD_AXIS_INDEX.horizontal, 'positive'),
    ];
  }
  if (fallback === 'vertical') {
    return [
      buttonKey(DPAD_BUTTON_INDEX.up),
      buttonKey(DPAD_BUTTON_INDEX.down),
      axisKey(DPAD_AXIS_INDEX.vertical, 'negative'),
      axisKey(DPAD_AXIS_INDEX.vertical, 'positive'),
    ];
  }
  return [];
}

/**
 * Warnings for the main Alt+C list, keyed by row (the action id, or its D-pad row key):
 * a row warns when another row uses one of its inputs. Every such pair fires together at
 * some point, since the situations nest (always, then keyboard-only mode, then popup open),
 * so a shared input is always worth pointing out. Dialog layouts are checked separately:
 * they pause these actions while open.
 */
export function findControllerBindingIssues(
  definitions: readonly ControllerBindingIssueDefinition[],
  bindings: ControllerBindings,
): Partial<Record<string, string>> {
  const issues: Partial<Record<string, string>> = {};
  const holdActions = new Map<string, string>();
  const rows: { key: string; label: string; inputs: string[] }[] = [];
  for (const definition of definitions) {
    const binding = bindings[definition.id];
    if (!binding) continue;
    if (definition.bindingType === 'axis') {
      const axisBinding = binding as ResolvedControllerAxisBinding;
      rows.push({
        key: definition.id,
        label: `${definition.label} (Stick)`,
        inputs: stickInputs(axisBinding),
      });
      rows.push({
        key: getControllerDpadRowKey(definition.id),
        label: `${definition.label} (D-pad)`,
        inputs: dpadInputs(axisBinding.kind === 'none' ? 'none' : axisBinding.dpadFallback),
      });
    } else {
      const discreteBinding = binding as ResolvedControllerDiscreteBinding;
      rows.push({
        key: definition.id,
        label: definition.label,
        inputs: discreteInputs(discreteBinding),
      });
      if (discreteBinding.kind !== 'none' && discreteBinding.hold) {
        const holdKey = controllerReviewInputSourceKey(discreteBinding.hold);
        if (holdKey === controllerReviewInputSourceKey(discreteBinding)) {
          issues[definition.id] = 'Hold and press are the same button';
        } else if (!holdActions.has(holdKey)) {
          holdActions.set(holdKey, definition.label);
        }
      }
    }
  }

  // Everything else on a hold button can wait for release. This one shows while held.
  const secondaryHold = bindings.holdSecondarySub;
  if (secondaryHold.kind !== 'none' && !secondaryHold.hold) {
    const comboLabel = holdActions.get(controllerReviewInputSourceKey(secondaryHold));
    if (comboLabel) issues.holdSecondarySub = `Also the hold button of ${comboLabel}`;
  }

  for (const row of rows) {
    const inputs = new Set(row.inputs);
    const clashes = rows
      .filter((other) => other !== row && other.inputs.some((input) => inputs.has(input)))
      .map((other) => other.label);
    if (clashes.length > 0) issues[row.key] = `Same as ${clashes.join(', ')}`;
  }
  return issues;
}
