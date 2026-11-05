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

/**
 * Rows that do nothing in SubMiner: no action reads `rightStickHorizontal`, so it cannot
 * clash with anything.
 */
const INACTIVE_BINDINGS: ReadonlySet<ControllerBindingActionId> = new Set(['rightStickHorizontal']);

/** Row key of a stick binding's D-pad row in Alt+C. */
export function getControllerDpadRowKey(actionId: ControllerBindingActionId): string {
  return `${actionId}:dpad`;
}

const buttonKey = (buttonIndex: number) => `button:${buttonIndex}`;
const axisKey = (axisIndex: number, direction: 'negative' | 'positive') =>
  `axis:${axisIndex}:${direction}`;

function discreteInputs(binding: ResolvedControllerDiscreteBinding): string[] {
  if (binding.kind === 'none') return [];
  if ('buttonIndex' in binding) return [buttonKey(binding.buttonIndex)];
  return [axisKey(binding.axisIndex, binding.direction)];
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
  const rows: { key: string; label: string; inputs: string[] }[] = [];
  for (const definition of definitions) {
    if (INACTIVE_BINDINGS.has(definition.id)) continue;
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
      rows.push({
        key: definition.id,
        label: definition.label,
        inputs: discreteInputs(binding as ResolvedControllerDiscreteBinding),
      });
    }
  }

  const issues: Partial<Record<string, string>> = {};
  for (const row of rows) {
    const inputs = new Set(row.inputs);
    const clashes = rows
      .filter((other) => other !== row && other.inputs.some((input) => inputs.has(input)))
      .map((other) => other.label);
    if (clashes.length > 0) issues[row.key] = `Same as ${clashes.join(', ')}`;
  }
  return issues;
}
