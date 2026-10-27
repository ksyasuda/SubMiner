import assert from 'node:assert/strict';
import test from 'node:test';

import { CORE_DEFAULT_CONFIG } from '../../config/definitions/defaults-core';
import {
  CONTROLLER_FIELD_GROUPING_ACTIONS,
  CONTROLLER_FIELD_GROUPING_SCREEN,
  type ControllerFieldGroupingStep,
} from '../../shared/controller-field-grouping';
import { findControllerScreenBindingIssues } from '../../shared/controller-review';
import type { ControllerFieldGroupingAction, ResolvedControllerConfig } from '../../types';
import {
  createControllerComboScreenLayout,
  formatControllerScreenHints,
  getControllerReviewInputContext,
} from './controller-screen-layout.js';
import type { GamepadLike } from './gamepad-controller.js';
import { resolveControllerConfigForGamepad } from '../controller-profile-config.js';

const CONFIG: ResolvedControllerConfig = { ...CORE_DEFAULT_CONFIG.controller, enabled: true };
const { repeatDelayMs, repeatIntervalMs } = CONFIG;

// Standard-layout indices; faces follow config.buttonIndices (Xbox positions by default).
const BUTTON = { A: 0, B: 1, X: 2, Y: 3, UP: 12, DOWN: 13, LEFT: 14, RIGHT: 15 };
type ButtonName = keyof typeof BUTTON;

function pad(pressed: ButtonName[] = []): GamepadLike {
  const buttons = Array.from({ length: 17 }, () => ({ value: 0, pressed: false }));
  for (const name of pressed) buttons[BUTTON[name]] = { value: 1, pressed: true };
  return { id: 'pad', index: 0, connected: true, mapping: 'standard', axes: [0, 0, 0, 0], buttons };
}

/** The field grouping layout driving a fake dialog that follows the default flow. */
function createFieldGroupingLayout(config = CONFIG) {
  const actions: ControllerFieldGroupingAction[] = [];
  const state: { step: ControllerFieldGroupingStep } = { step: 'select' };
  const layout = createControllerComboScreenLayout({
    screen: CONTROLLER_FIELD_GROUPING_SCREEN,
    getBindings: (resolved) => resolved.fieldGroupingBindings,
    getStep: () => state.step,
    dispatch: (action) => {
      actions.push(action);
      if (action === 'continue') state.step = 'preview';
      if (action === 'back') state.step = 'select';
    },
  });
  layout.activate(pad(), config, 0, { latchHeld: true });
  return { layout, actions, state };
}

test('the default duplicate card layout binds every action but preview Cancel, without conflicts', () => {
  for (const action of CONTROLLER_FIELD_GROUPING_ACTIONS) {
    assert.equal(
      CONFIG.fieldGroupingBindings[action.id]?.kind,
      action.id === 'cancelMerge' ? 'none' : 'combo',
      action.id,
    );
  }
  for (const mapping of ['standard', ''] as const) {
    for (const faceButtonLabels of ['xbox', 'nintendo'] as const) {
      const context = getControllerReviewInputContext({ ...CONFIG, faceButtonLabels }, mapping);
      assert.deepEqual(
        findControllerScreenBindingIssues(
          CONTROLLER_FIELD_GROUPING_SCREEN,
          CONFIG.fieldGroupingBindings,
          context,
        ),
        {},
      );
    }
  }
});

test('duplicate card hints read like the agreed layout', () => {
  const hints = formatControllerScreenHints(
    CONTROLLER_FIELD_GROUPING_SCREEN,
    CONFIG.fieldGroupingBindings,
    CONFIG,
    'standard',
  );
  assert.deepEqual(hints, {
    selectOriginal: '◀',
    selectNew: '▶',
    toggleDeleteDuplicate: 'X',
    continue: 'A',
    cancel: 'B',
    confirmMerge: 'A',
    back: 'B',
    cancelMerge: null,
    toggleView: 'Y',
    scrollUp: '▲',
    scrollDown: '▼',
  });
});

test('the D-pad picks the card and X toggles delete on the first step', () => {
  const { layout, actions } = createFieldGroupingLayout();

  layout.poll(pad(['RIGHT']), CONFIG, 10);
  layout.poll(pad(['RIGHT']), CONFIG, 2_000);
  layout.poll(pad(), CONFIG, 2_010);
  layout.poll(pad(['LEFT']), CONFIG, 2_020);
  layout.poll(pad(['X']), CONFIG, 2_030);
  // Up/down scroll only on the preview.
  layout.poll(pad(['DOWN']), CONFIG, 2_040);

  assert.deepEqual(actions, ['selectNew', 'selectOriginal', 'toggleDeleteDuplicate']);
});

test('B on the preview goes back; holding it does not also cancel', () => {
  const { layout, actions, state } = createFieldGroupingLayout();

  layout.poll(pad(['A']), CONFIG, 10);
  assert.equal(state.step, 'preview');
  // A still held after the step change does not confirm the merge.
  layout.poll(pad(['A']), CONFIG, 20);
  layout.poll(pad(), CONFIG, 30);
  layout.poll(pad(['B']), CONFIG, 40);
  assert.equal(state.step, 'select');
  // B held across the step change is not a fresh press on the first step.
  layout.poll(pad(['B']), CONFIG, 1_000);
  layout.poll(pad(), CONFIG, 1_010);
  layout.poll(pad(['B']), CONFIG, 1_020);

  assert.deepEqual(actions, ['continue', 'back', 'cancel']);
});

test('preview scrolling repeats while the D-pad is held', () => {
  const { layout, actions, state } = createFieldGroupingLayout();
  state.step = 'preview';

  layout.poll(pad(['DOWN']), CONFIG, 0);
  layout.poll(pad(['DOWN']), CONFIG, repeatDelayMs - 1);
  layout.poll(pad(['DOWN']), CONFIG, repeatDelayMs);
  layout.poll(pad(['DOWN']), CONFIG, repeatDelayMs + repeatIntervalMs);
  layout.poll(pad(), CONFIG, repeatDelayMs + repeatIntervalMs + 1);
  layout.poll(pad(['Y']), CONFIG, 5_000);
  layout.poll(pad(['Y']), CONFIG, 6_000);

  assert.deepEqual(actions, ['scrollDown', 'scrollDown', 'scrollDown', 'toggleView']);
});

test('remapped duplicate card bindings replace the defaults', () => {
  const config: ResolvedControllerConfig = {
    ...CONFIG,
    fieldGroupingBindings: {
      ...CONFIG.fieldGroupingBindings,
      back: { kind: 'none' },
      cancelMerge: { kind: 'combo', press: { kind: 'face', letter: 'b' }, hold: null },
    },
  };
  const { layout, actions, state } = createFieldGroupingLayout(config);
  state.step = 'preview';

  layout.poll(pad(['B']), config, 10);

  assert.deepEqual(actions, ['cancelMerge']);
});

test('dialog hints follow the button style of the controller in use', () => {
  const config: ResolvedControllerConfig = {
    ...CONFIG,
    profiles: {
      'pad-ps': {
        label: 'DualSense',
        faceButtonLabels: 'playstation',
        buttonIndices: CONFIG.buttonIndices,
        bindings: CONFIG.bindings,
        reviewBindings: CONFIG.reviewBindings,
        fieldGroupingBindings: CONFIG.fieldGroupingBindings,
      },
    },
  };
  const hintsFor = (gamepadId: string) => {
    const padConfig = resolveControllerConfigForGamepad(config, gamepadId);
    return formatControllerScreenHints(
      CONTROLLER_FIELD_GROUPING_SCREEN,
      padConfig.fieldGroupingBindings,
      padConfig,
      'standard',
    );
  };

  assert.equal(hintsFor('pad-ps').continue, '×');
  assert.equal(hintsFor('pad-ps').toggleView, '△');
  // Controllers without their own style use controller.faceButtonLabels.
  assert.equal(hintsFor('other-pad').continue, 'A');
});
