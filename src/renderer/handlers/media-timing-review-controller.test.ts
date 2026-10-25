import assert from 'node:assert/strict';
import test from 'node:test';

import { CORE_DEFAULT_CONFIG } from '../../config/definitions/defaults-core';
import type { ResolvedControllerConfig } from '../../types';
import type { GamepadLike } from './gamepad-controller.js';
import {
  MEDIA_TIMING_REVIEW_CONTROLLER_LAYOUT,
  createMediaTimingReviewControllerLayout,
  formatMediaTimingReviewControllerHint,
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
  layout.activate(pad(), CONFIG, 0);
  return { layout, actions, state };
}

test('every review action has exactly one controller binding per step', () => {
  const seenInputs = new Set<string>();
  const seenActions = new Set<string>();
  for (const binding of MEDIA_TIMING_REVIEW_CONTROLLER_LAYOUT) {
    const input = `${binding.step}:${formatMediaTimingReviewControllerHint(binding)}`;
    assert.equal(seenInputs.has(input), false, `duplicate input ${input}`);
    assert.equal(seenActions.has(binding.action), false, `duplicate action ${binding.action}`);
    seenInputs.add(input);
    seenActions.add(binding.action);
  }
  assert.equal(seenActions.size, 21);
});

test('controller hints read like the documented layout', () => {
  const hint = (action: MediaTimingReviewControllerAction) => {
    const binding = MEDIA_TIMING_REVIEW_CONTROLLER_LAYOUT.find((entry) => entry.action === action);
    assert.ok(binding);
    return formatMediaTimingReviewControllerHint(binding);
  };
  assert.equal(hint('confirm'), 'A');
  assert.equal(hint('showEarlier'), '◀');
  assert.equal(hint('startForward'), 'L1 + ▶');
  assert.equal(hint('nextLineAdd'), 'R1 + ▲');
  assert.equal(hint('frameReset'), 'L1 + R1');
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
  layout.activate(pad(['X', 'RIGHT']), CONFIG, 0);
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
