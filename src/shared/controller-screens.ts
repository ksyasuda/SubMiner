import { CONTROLLER_FIELD_GROUPING_SCREEN } from './controller-field-grouping';
import {
  CONTROLLER_REVIEW_SCREEN,
  isControllerScreenAction,
  parseControllerReviewBinding,
  type ControllerScreenBindings,
  type ControllerScreenDefinition,
} from './controller-review';

/** Every dialog with its own remappable controller layout, in Alt+C order. */
export const CONTROLLER_SCREENS: readonly ControllerScreenDefinition[] = [
  CONTROLLER_REVIEW_SCREEN as unknown as ControllerScreenDefinition,
  CONTROLLER_FIELD_GROUPING_SCREEN as unknown as ControllerScreenDefinition,
];

export type ControllerScreenId = ControllerScreenDefinition['id'];
export type ControllerScreenBindingsKey = ControllerScreenDefinition['configKey'];

export function getControllerScreen(id: ControllerScreenId): ControllerScreenDefinition {
  return CONTROLLER_SCREENS.find((screen) => screen.id === id)!;
}

/**
 * Parses a per-action patch for one screen (`{ "continue": { … } }`). Returns null when any
 * key is not an action of that screen or any binding is malformed.
 */
export function parseControllerScreenBindingsPatch(
  screen: ControllerScreenDefinition,
  value: unknown,
): Partial<ControllerScreenBindings> | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return null;
  const bindings: Partial<ControllerScreenBindings> = {};
  for (const [action, rawBinding] of Object.entries(value)) {
    if (!isControllerScreenAction(screen, action)) return null;
    const parsed = parseControllerReviewBinding(rawBinding);
    if (!parsed) return null;
    bindings[action] = parsed;
  }
  return bindings;
}
