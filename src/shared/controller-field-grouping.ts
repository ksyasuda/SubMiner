import type {
  ControllerFieldGroupingAction,
  ResolvedControllerFieldGroupingBindings,
} from '../types/runtime';
import {
  combo,
  dpad,
  face,
  type ControllerScreenActionDefinition,
  type ControllerScreenDefinition,
} from './controller-review';

/** The duplicate-card dialog: pick the card to keep, then the final merge preview. */
export type ControllerFieldGroupingStep = 'select' | 'preview';

export type ControllerFieldGroupingActionDefinition = ControllerScreenActionDefinition<
  ControllerFieldGroupingAction,
  ControllerFieldGroupingStep
>;

/** Every field grouping action, in the order Alt+C lists them. */
export const CONTROLLER_FIELD_GROUPING_ACTIONS: readonly ControllerFieldGroupingActionDefinition[] =
  [
    { id: 'selectOriginal', label: 'Original card', step: 'select', repeat: false },
    { id: 'selectNew', label: 'New card', step: 'select', repeat: false },
    {
      id: 'toggleDeleteDuplicate',
      label: 'Delete duplicate on/off',
      step: 'select',
      repeat: false,
    },
    { id: 'continue', label: 'Continue', step: 'select', repeat: false },
    { id: 'cancel', label: 'Cancel', step: 'select', repeat: false },
    { id: 'confirmMerge', label: 'Confirm merge', step: 'preview', repeat: false },
    { id: 'back', label: 'Back', step: 'preview', repeat: false },
    { id: 'cancelMerge', label: 'Cancel', step: 'preview', repeat: false },
    { id: 'toggleView', label: 'Compact / Full', step: 'preview', repeat: false },
    { id: 'scrollUp', label: 'Scroll up', step: 'preview', repeat: true },
    { id: 'scrollDown', label: 'Scroll down', step: 'preview', repeat: true },
  ];

/**
 * Default layout: D-pad picks the card, A goes forward, B goes back (B on the preview
 * returns to the card choice, B there cancels). Cancel on the preview is unbound, so a
 * double B press is the way out unless the user binds it.
 */
const DEFAULT_CONTROLLER_FIELD_GROUPING_BINDINGS: ResolvedControllerFieldGroupingBindings = {
  selectOriginal: combo(dpad('left')),
  selectNew: combo(dpad('right')),
  toggleDeleteDuplicate: combo(face('x')),
  continue: combo(face('a')),
  cancel: combo(face('b')),
  confirmMerge: combo(face('a')),
  back: combo(face('b')),
  cancelMerge: { kind: 'none' },
  toggleView: combo(face('y')),
  scrollUp: combo(dpad('up')),
  scrollDown: combo(dpad('down')),
};

export const CONTROLLER_FIELD_GROUPING_SCREEN: ControllerScreenDefinition<
  ControllerFieldGroupingAction,
  ControllerFieldGroupingStep
> = {
  id: 'fieldGrouping',
  title: 'Duplicate card (field grouping)',
  configKey: 'fieldGroupingBindings',
  describe: 'Duplicate card dialog',
  steps: [
    { id: 'select', label: 'Choosing the card' },
    { id: 'preview', label: 'Merge preview' },
  ],
  actions: CONTROLLER_FIELD_GROUPING_ACTIONS,
  defaults: DEFAULT_CONTROLLER_FIELD_GROUPING_BINDINGS,
};

export function getDefaultControllerFieldGroupingBindings(): ResolvedControllerFieldGroupingBindings {
  return structuredClone(DEFAULT_CONTROLLER_FIELD_GROUPING_BINDINGS);
}
