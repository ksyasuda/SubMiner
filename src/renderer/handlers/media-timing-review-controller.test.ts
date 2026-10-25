import assert from 'node:assert/strict';
import test from 'node:test';

import { CORE_DEFAULT_CONFIG } from '../../config/definitions/defaults-core';
import {
  CONTROLLER_REVIEW_ACTIONS,
  findControllerReviewBindingIssues,
} from '../../shared/controller-review';
import type { ResolvedControllerConfig } from '../../types';
import type { GamepadLike } from './gamepad-controller.js';
import {
  createMediaTimingReviewControllerLayout,
  formatMediaTimingReviewControllerHints,
  getControllerReviewInputContext,
  type MediaTimingReviewControllerAction,
  type MediaTimingReviewStep,
} from './media-timing-review-controller.js';

const CONFIG: ResolvedControllerConfig = { ...CORE_DEFAULT_CONFIG.controller, enabled: true };
const { repeatDelayMs, repeatIntervalMs } = CONFIG;

// Standard-layout indices; the layout reads them from config.buttonIndices.
const BUTTON = { A: 0, B: 1, X: 2, Y: 3, L1: 4, R1: 5, UP: 12, DOWN: 13, LEFT: 14, RIGHT: 15 };
type ButtonName = keyof typeof BUTTON;

function pad(pressed: ButtonName[] = [], axes: number[] = [0, 0, 0, 0]): GamepadLike {
  const buttons = Array.from({ length: 17 }, () => ({ value: 0, pressed: false }));
  for (const name of pressed) buttons[BUTTON[name]] = { value: 1, pressed: true };
  return { id: 'pad', index: 0, connected: true, mapping: 'standard', axes, buttons };
}

function createLayout(step: MediaTimingReviewStep = 'editor') {
  const actions: MediaTimingReviewControllerAction[] = [];
  const state = { step };
  const layout = createMediaTimingReviewControllerLayout({
    getStep: () => state.step,
    dispatch: (action) => actions.push(action),
  });
  layout.activate(pad(), CONFIG, 0, { latchHeld: true });
  return { layout, actions, state };
}

test('the default review layout binds all 21 actions without conflicts', () => {
  assert.equal(CONTROLLER_REVIEW_ACTIONS.length, 21);
  for (const action of CONTROLLER_REVIEW_ACTIONS) {
    assert.equal(CONFIG.reviewBindings[action.id]?.kind, 'combo', action.id);
  }
  for (const mapping of ['standard', ''] as const) {
    for (const faceButtonLabels of ['xbox', 'nintendo'] as const) {
      const context = getControllerReviewInputContext({ ...CONFIG, faceButtonLabels }, mapping);
      assert.deepEqual(findControllerReviewBindingIssues(CONFIG.reviewBindings, context), {});
    }
  }
});

test('controller hints read like the documented layout', () => {
  const hints = formatMediaTimingReviewControllerHints(CONFIG, 'standard');
  assert.equal(hints.confirm, 'A');
  assert.equal(hints.showEarlier, '\u25c0');
  assert.equal(hints.startForward, 'L1 + \u25b6');
  assert.equal(hints.nextLineAdd, 'R1 + \u25b2');
  assert.equal(hints.frameReset, 'L1 + R1');
});

test('face buttons press the editor buttons once per press', () => {
  const { layout, actions } = createLayout();

  layout.poll(pad(['A']), CONFIG, 10);
  layout.poll(pad(['A']), CONFIG, 900);
  layout.poll(pad(), CONFIG, 910);
  layout.poll(pad(['B', 'X', 'Y']), CONFIG, 920);

  assert.deepEqual(actions, ['confirm', 'cancel', 'play', 'resetTiming']);
});

test('face buttons press the cancel-step buttons on that step', () => {
  const { layout, actions } = createLayout('cancel');

  layout.poll(pad(['A', 'B', 'X', 'Y']), CONFIG, 10);

  assert.deepEqual(actions, ['useOriginal', 'keepEditing', 'skipMedia', 'discard']);
});

test('buttons already held when the review opens wait for a fresh press', () => {
  const actions: MediaTimingReviewControllerAction[] = [];
  const layout = createMediaTimingReviewControllerLayout({
    getStep: () => 'editor',
    dispatch: (action) => actions.push(action),
  });

  // X mined the card that opened this review; it must not also start the preview.
  layout.activate(pad(['X', 'RIGHT']), CONFIG, 0, { latchHeld: true });
  layout.poll(pad(['X', 'RIGHT']), CONFIG, 1_000);
  assert.deepEqual(actions, []);

  layout.poll(pad(), CONFIG, 1_010);
  layout.poll(pad(['X', 'RIGHT']), CONFIG, 1_020);
  assert.deepEqual(actions, ['play', 'showLater']);
});

test('D-pad alone moves the timeline and steps screenshot frames', () => {
  const { layout, actions } = createLayout();

  for (const name of ['LEFT', 'RIGHT', 'UP', 'DOWN'] as const) {
    layout.poll(pad([name]), CONFIG, 10);
    layout.poll(pad(), CONFIG, 20);
  }

  assert.deepEqual(actions, ['showEarlier', 'showLater', 'framePrevious', 'frameNext']);
});

test('L1 combos adjust the start and previous lines, R1 combos the end and next lines', () => {
  const { layout, actions } = createLayout();
  const press = (names: ButtonName[]) => {
    layout.poll(pad(names), CONFIG, 10);
    layout.poll(pad(), CONFIG, 20);
  };

  press(['L1', 'LEFT']);
  press(['L1', 'RIGHT']);
  press(['L1', 'UP']);
  press(['L1', 'DOWN']);
  press(['R1', 'LEFT']);
  press(['R1', 'RIGHT']);
  press(['R1', 'UP']);
  press(['R1', 'DOWN']);

  assert.deepEqual(actions, [
    'startBack',
    'startForward',
    'previousLineAdd',
    'previousLineRemove',
    'endBack',
    'endForward',
    'nextLineAdd',
    'nextLineRemove',
  ]);
});

test('holding a D-pad combo repeats after the configured delay', () => {
  const { layout, actions } = createLayout();
  const hold = (now: number) => layout.poll(pad(['L1', 'RIGHT']), CONFIG, now);

  hold(1_000);
  hold(1_000 + repeatDelayMs - 1);
  hold(1_000 + repeatDelayMs);
  hold(1_000 + repeatDelayMs + repeatIntervalMs - 1);
  hold(1_000 + repeatDelayMs + repeatIntervalMs);

  assert.deepEqual(actions, ['startForward', 'startForward', 'startForward']);
});

test('letting go of L1 before the D-pad does not fire the plain D-pad action', () => {
  const { layout, actions } = createLayout();

  layout.poll(pad(['L1', 'RIGHT']), CONFIG, 10);
  layout.poll(pad(['RIGHT']), CONFIG, 20);
  layout.poll(pad(['RIGHT']), CONFIG, 2_000);
  assert.deepEqual(actions, ['startForward']);

  layout.poll(pad(), CONFIG, 2_010);
  layout.poll(pad(['RIGHT']), CONFIG, 2_020);
  assert.deepEqual(actions, ['startForward', 'showLater']);
});

test('pressing L1 and R1 together resets the screenshot frame once', () => {
  const { layout, actions } = createLayout();

  layout.poll(pad(['L1']), CONFIG, 10);
  layout.poll(pad(['L1', 'R1']), CONFIG, 20);
  layout.poll(pad(['L1', 'R1', 'RIGHT']), CONFIG, 30);
  layout.poll(pad(['L1', 'R1', 'RIGHT']), CONFIG, 2_000);

  assert.deepEqual(actions, ['frameReset']);
});

test('D-pads reported as axes (non-standard pads) drive the same layout', () => {
  const { layout, actions } = createLayout();
  const axes = (horizontal: number, vertical: number) => [0, 0, 0, 0, 0, 0, horizontal, vertical];

  layout.poll(pad(['R1'], axes(-1, 0)), CONFIG, 10);
  layout.poll(pad([], axes(0, 0)), CONFIG, 20);
  layout.poll(pad([], axes(0, 1)), CONFIG, 30);

  assert.deepEqual(actions, ['endBack', 'frameNext']);
});

test('Nintendo-style pads press the button labelled A to confirm, wherever it sits', () => {
  const nintendo: ResolvedControllerConfig = { ...CONFIG, faceButtonLabels: 'nintendo' };
  const actions: MediaTimingReviewControllerAction[] = [];
  const layout = createMediaTimingReviewControllerLayout({
    getStep: () => 'editor',
    dispatch: (action) => actions.push(action),
  });
  layout.activate(pad(), nintendo, 0, { latchHeld: true });

  // Positions, not letters: on these pads A is the right button (index 1), B the bottom
  // one (0), X the top one (3) and Y the left one (2).
  for (const position of ['B', 'A', 'Y', 'X'] as const) {
    layout.poll(pad([position]), nintendo, 10);
    layout.poll(pad(), nintendo, 20);
  }

  assert.deepEqual(actions, ['confirm', 'cancel', 'play', 'resetTiming']);
});

test('a press that first reveals the pad to a fresh dialog window counts', () => {
  const actions: MediaTimingReviewControllerAction[] = [];
  const layout = createMediaTimingReviewControllerLayout({
    getStep: () => 'editor',
    dispatch: (action) => actions.push(action),
  });

  layout.activate(pad(['A']), CONFIG, 0, { latchHeld: false });
  layout.poll(pad(['A']), CONFIG, 16);

  assert.deepEqual(actions, ['confirm']);
});

function withReviewBindings(
  overrides: Partial<ResolvedControllerConfig['reviewBindings']>,
): ResolvedControllerConfig {
  return { ...CONFIG, reviewBindings: { ...CONFIG.reviewBindings, ...overrides } };
}

test('remapped combos use any hold button and leave plain presses alone', () => {
  const config = withReviewBindings({
    // Hold L2 (raw button 6 on standard pads) and press A.
    startForward: {
      kind: 'combo',
      press: { kind: 'face', letter: 'a' },
      hold: { kind: 'button', buttonIndex: 6 },
    },
  });
  const actions: MediaTimingReviewControllerAction[] = [];
  const layout = createMediaTimingReviewControllerLayout({
    getStep: () => 'editor',
    dispatch: (action) => actions.push(action),
  });
  const press = (indices: number[], now: number) => {
    const buttons = Array.from({ length: 17 }, (_, index) => ({
      value: indices.includes(index) ? 1 : 0,
      pressed: indices.includes(index),
    }));
    layout.poll({ ...pad(), buttons }, config, now);
  };
  layout.activate(pad(), config, 0, { latchHeld: true });

  press([6], 10);
  press([6, 0], 20);
  press([], 30);
  press([0], 40);

  assert.deepEqual(actions, ['startForward', 'confirm']);
});

test('unbound review actions never fire', () => {
  const config = withReviewBindings({ confirm: { kind: 'none' } });
  const actions: MediaTimingReviewControllerAction[] = [];
  const layout = createMediaTimingReviewControllerLayout({
    getStep: () => 'editor',
    dispatch: (action) => actions.push(action),
  });
  layout.activate(pad(), config, 0, { latchHeld: true });

  layout.poll(pad(['A']), config, 10);

  assert.deepEqual(actions, []);
});

test('repeat follows the action, not the button it is on', () => {
  const config = withReviewBindings({
    startForward: { kind: 'combo', press: { kind: 'face', letter: 'b' }, hold: null },
    cancel: { kind: 'combo', press: { kind: 'dpad', direction: 'right' }, hold: null },
    showLater: { kind: 'none' },
  });
  const actions: MediaTimingReviewControllerAction[] = [];
  const layout = createMediaTimingReviewControllerLayout({
    getStep: () => 'editor',
    dispatch: (action) => actions.push(action),
  });
  layout.activate(pad(), config, 0, { latchHeld: true });

  for (const now of [10, 10 + repeatDelayMs, 10 + repeatDelayMs + repeatIntervalMs]) {
    layout.poll(pad(['B', 'RIGHT']), config, now);
  }

  // Both fire on the first press; only startForward repeats.
  assert.deepEqual(actions, ['cancel', 'startForward', 'startForward', 'startForward']);
});

test('conflicts are reported, including a learned D-pad button that matches a default', () => {
  const config = withReviewBindings({
    // A learned raw D-pad left (button 14) is the same input as the default "D-pad ◀".
    showLater: { kind: 'combo', press: { kind: 'button', buttonIndex: 14 }, hold: null },
    // R1 alone, while other actions hold R1.
    play: { kind: 'combo', press: { kind: 'button', buttonIndex: 5 }, hold: null },
  });
  const issues = findControllerReviewBindingIssues(
    config.reviewBindings,
    getControllerReviewInputContext(config, 'standard'),
  );

  assert.equal(issues.showLater, 'Same as Earlier −2s');
  assert.equal(issues.showEarlier, 'Same as Later +2s');
  assert.equal(issues.play, 'R1 is also a hold button on this step');
});
