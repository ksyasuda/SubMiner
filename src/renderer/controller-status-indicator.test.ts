import assert from 'node:assert/strict';
import test from 'node:test';

import { createControllerStatusIndicator } from './controller-status-indicator.js';

const pad = (id: string, index = 0) => ({ id, index, mapping: 'standard', connected: true });

test('controller status indicator notifies once when a controller is first detected', () => {
  const messages: string[] = [];
  const indicator = createControllerStatusIndicator((message) => messages.push(message));

  indicator.update({ connectedGamepads: [], activeGamepadId: null });
  assert.deepEqual(messages, []);

  indicator.update({ connectedGamepads: [pad('pad-1')], activeGamepadId: 'pad-1' });
  indicator.update({ connectedGamepads: [pad('pad-1')], activeGamepadId: 'pad-1' });
  assert.deepEqual(messages, ['Controller connected: pad-1']);
});

test('controller status indicator announces newly detected controllers after startup', () => {
  const messages: string[] = [];
  const indicator = createControllerStatusIndicator((message) => messages.push(message));

  indicator.update({ connectedGamepads: [pad('pad-1')], activeGamepadId: 'pad-1' });
  indicator.update({
    connectedGamepads: [pad('pad-1'), pad('pad-2', 1)],
    activeGamepadId: 'pad-1',
  });

  assert.deepEqual(messages, ['Controller connected: pad-1', 'Controller connected: pad-2']);
});

test('controller status indicator announces a controller again after it reconnects', () => {
  const messages: string[] = [];
  const indicator = createControllerStatusIndicator((message) => messages.push(message));

  indicator.update({ connectedGamepads: [pad('pad-1')], activeGamepadId: 'pad-1' });
  indicator.update({ connectedGamepads: [], activeGamepadId: null });
  indicator.update({ connectedGamepads: [pad('pad-1')], activeGamepadId: 'pad-1' });

  assert.equal(messages.length, 2);
});
