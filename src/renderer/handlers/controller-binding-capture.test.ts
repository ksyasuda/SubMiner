import assert from 'node:assert/strict';
import test from 'node:test';

import { createControllerBindingCapture } from './controller-binding-capture.js';

function createSnapshot(
  overrides: {
    axes?: number[];
    buttons?: Array<{ value: number; pressed?: boolean; touched?: boolean }>;
  } = {},
) {
  return {
    axes: overrides.axes ?? [0, 0, 0, 0, 0],
    buttons:
      overrides.buttons ??
      Array.from({ length: 12 }, () => ({
        value: 0,
        pressed: false,
        touched: false,
      })),
  };
}

test('controller binding capture waits for neutral-to-active button edge', () => {
  const capture = createControllerBindingCapture({
    triggerDeadzone: 0.5,
    stickDeadzone: 0.2,
  });

  const heldButtons = createSnapshot({
    buttons: [{ value: 1, pressed: true, touched: true }],
  });

  capture.arm({ actionId: 'toggleLookup', bindingType: 'discrete' }, heldButtons);

  assert.equal(capture.poll(heldButtons), null);

  const neutralButtons = createSnapshot();
  assert.equal(capture.poll(neutralButtons), null);

  const freshPress = createSnapshot({
    buttons: [{ value: 1, pressed: true, touched: true }],
  });
  assert.deepEqual(capture.poll(freshPress), {
    actionId: 'toggleLookup',
    bindingType: 'discrete',
    binding: { kind: 'button', buttonIndex: 0 },
  });
});

test('controller binding capture records fresh axis direction for discrete learn mode', () => {
  const capture = createControllerBindingCapture({
    triggerDeadzone: 0.5,
    stickDeadzone: 0.2,
  });

  capture.arm({ actionId: 'closeLookup', bindingType: 'discrete' }, createSnapshot());

  assert.deepEqual(capture.poll(createSnapshot({ axes: [0, 0, 0, -0.8] })), {
    actionId: 'closeLookup',
    bindingType: 'discrete',
    binding: { kind: 'axis', axisIndex: 3, direction: 'negative' },
  });
});

test('controller binding capture ignores analog drift inside deadzone', () => {
  const capture = createControllerBindingCapture({
    triggerDeadzone: 0.5,
    stickDeadzone: 0.3,
  });

  capture.arm({ actionId: 'mineCard', bindingType: 'discrete' }, createSnapshot());

  assert.equal(capture.poll(createSnapshot({ axes: [0.2, 0, 0, 0] })), null);
  assert.equal(capture.isArmed(), true);
});

test('controller binding capture emits axis binding for continuous learn mode', () => {
  const capture = createControllerBindingCapture({
    triggerDeadzone: 0.5,
    stickDeadzone: 0.2,
  });

  capture.arm(
    {
      actionId: 'leftStickHorizontal',
      bindingType: 'axis',
      dpadFallback: 'horizontal',
    },
    createSnapshot(),
  );

  assert.deepEqual(capture.poll(createSnapshot({ axes: [0, 0, 0, 0.9] })), {
    actionId: 'leftStickHorizontal',
    bindingType: 'axis',
    binding: { kind: 'axis', axisIndex: 3, dpadFallback: 'horizontal' },
  });
});
