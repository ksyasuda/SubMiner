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

test('controller binding capture waits for a neutral-to-active D-pad edge', () => {
  const capture = createControllerBindingCapture({
    triggerDeadzone: 0.5,
    stickDeadzone: 0.2,
  });

  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  const heldDpad = createSnapshot({
    buttons: buttons.map((button, index) =>
      index === 15 ? { value: 1, pressed: true, touched: true } : button,
    ),
  });

  capture.arm({ actionId: 'leftStickHorizontal', bindingType: 'dpad' }, heldDpad);

  assert.equal(capture.poll(heldDpad), null);
  assert.equal(capture.poll(createSnapshot({ buttons })), null);
  assert.deepEqual(capture.poll(heldDpad), {
    actionId: 'leftStickHorizontal',
    bindingType: 'dpad',
    dpadDirection: 'horizontal',
  });
});

test('controller binding capture ignores analog drift inside deadzone', () => {
  const capture = createControllerBindingCapture({
    triggerDeadzone: 0.5,
    stickDeadzone: 0.3,
  });

  capture.arm(
    { actionId: 'leftStickHorizontal', bindingType: 'axis', dpadFallback: 'none' },
    createSnapshot(),
  );

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

test('controller binding capture ignores button presses for continuous learn mode', () => {
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

  assert.equal(
    capture.poll(
      createSnapshot({
        buttons: [{ value: 1, pressed: true, touched: true }],
      }),
    ),
    null,
  );

  assert.deepEqual(capture.poll(createSnapshot({ axes: [0, 0, 0.75, 0, 0] })), {
    actionId: 'leftStickHorizontal',
    bindingType: 'axis',
    binding: { kind: 'axis', axisIndex: 2, dpadFallback: 'horizontal' },
  });
});
