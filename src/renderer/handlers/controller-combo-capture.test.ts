import assert from 'node:assert/strict';
import test from 'node:test';

import { createControllerComboCapture } from './controller-combo-capture.js';

const PRESSED = { value: 1, pressed: true };
const RELEASED = { value: 0, pressed: false };

function snapshot(pressed: number[] = [], axes: number[] = [0, 0, 0, 0]) {
  return {
    axes,
    buttons: Array.from({ length: 17 }, (_, index) =>
      pressed.includes(index) ? PRESSED : RELEASED,
    ),
  };
}

function createCapture() {
  return createControllerComboCapture({ triggerDeadzone: 0.5, stickDeadzone: 0.2 });
}

test('combo capture records a single button once it is released', () => {
  const capture = createCapture();
  capture.arm(snapshot());

  assert.deepEqual(capture.poll(snapshot()), { status: 'waiting' });
  assert.deepEqual(capture.poll(snapshot([1])), {
    status: 'holding',
    hold: { kind: 'button', buttonIndex: 1 },
  });
  assert.deepEqual(capture.poll(snapshot()), {
    status: 'done',
    binding: { kind: 'combo', press: { kind: 'button', buttonIndex: 1 }, hold: null },
  });
});

test('combo capture records hold first + press second', () => {
  const capture = createCapture();
  capture.arm(snapshot());

  capture.poll(snapshot([4]));
  assert.deepEqual(capture.poll(snapshot([4, 15])), {
    status: 'done',
    binding: {
      kind: 'combo',
      press: { kind: 'button', buttonIndex: 15 },
      hold: { kind: 'button', buttonIndex: 4 },
    },
  });
});

test('combo capture ignores buttons held when Learn starts until they are released', () => {
  const capture = createCapture();
  capture.arm(snapshot([0]));

  assert.deepEqual(capture.poll(snapshot([0])), { status: 'waiting' });
  capture.poll(snapshot());
  assert.deepEqual(capture.poll(snapshot([0])), {
    status: 'holding',
    hold: { kind: 'button', buttonIndex: 0 },
  });
});

test('combo capture never records a trigger axis resting position', () => {
  const capture = createCapture();
  // Some pads rest trigger axes at -1.
  capture.arm(snapshot([], [0, 0, -1, 0]));

  capture.poll(snapshot([], [0, 0, -0.2, 0]));
  assert.deepEqual(capture.poll(snapshot([], [0, 0, -1, 0])), { status: 'waiting' });
  assert.deepEqual(capture.poll(snapshot([], [0, 0, 0.9, 0])), {
    status: 'holding',
    hold: { kind: 'axis', axisIndex: 2, direction: 'positive' },
  });
});

test('combo capture records one axis from a diagonal push, not a hold and a press', () => {
  const capture = createCapture();
  capture.arm(snapshot());
  const diagonal = snapshot([], [0.9, 0.9, 0, 0]);

  const hold = { kind: 'axis', axisIndex: 0, direction: 'positive' } as const;
  assert.deepEqual(capture.poll(diagonal), { status: 'holding', hold });
  assert.deepEqual(capture.poll(diagonal), { status: 'holding', hold });
  assert.deepEqual(capture.poll(snapshot()), {
    status: 'done',
    binding: { kind: 'combo', press: hold, hold: null },
  });
});

test('combo capture still records a second axis pressed after the first', () => {
  const capture = createCapture();
  capture.arm(snapshot());

  capture.poll(snapshot([], [0.9, 0, 0, 0]));
  assert.deepEqual(capture.poll(snapshot([], [0.9, 0.9, 0, 0])), {
    status: 'done',
    binding: {
      kind: 'combo',
      press: { kind: 'axis', axisIndex: 1, direction: 'positive' },
      hold: { kind: 'axis', axisIndex: 0, direction: 'positive' },
    },
  });
});

test('combo capture records a button and an axis activated in the same poll as a combo', () => {
  const capture = createCapture();
  capture.arm(snapshot());
  const both = snapshot([4], [0.9, 0, 0, 0]);

  const hold = { kind: 'button', buttonIndex: 4 } as const;
  assert.deepEqual(capture.poll(both), { status: 'holding', hold });
  assert.deepEqual(capture.poll(both), {
    status: 'done',
    binding: {
      kind: 'combo',
      press: { kind: 'axis', axisIndex: 0, direction: 'positive' },
      hold,
    },
  });
});
