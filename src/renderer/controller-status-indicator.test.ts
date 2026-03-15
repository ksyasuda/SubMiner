import assert from 'node:assert/strict';
import test from 'node:test';

import { createControllerStatusIndicator } from './controller-status-indicator.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

test('controller status indicator shows once when a controller is first detected and auto-hides', () => {
  let nextTimerId = 1;
  const scheduled = new Map<number, () => void>();
  const classList = createClassList(['hidden']);
  const toast = {
    textContent: '',
    classList,
  };

  const indicator = createControllerStatusIndicator({ controllerStatusToast: toast } as never, {
    durationMs: 1500,
    setTimeout: (callback: () => void) => {
      const id = nextTimerId++;
      scheduled.set(id, callback);
      return id as never;
    },
    clearTimeout: (id) => {
      scheduled.delete(id as never as number);
    },
  });

  indicator.update({
    connectedGamepads: [],
    activeGamepadId: null,
  });

  assert.equal(classList.contains('hidden'), true);
  assert.equal(toast.textContent, '');

  indicator.update({
    connectedGamepads: [{ id: 'pad-1', index: 0, mapping: 'standard', connected: true }],
    activeGamepadId: 'pad-1',
  });

  assert.equal(classList.contains('hidden'), false);
  assert.match(toast.textContent, /controller detected/i);
  assert.match(toast.textContent, /pad-1/i);
  assert.equal(scheduled.size, 1);

  indicator.update({
    connectedGamepads: [{ id: 'pad-1', index: 0, mapping: 'standard', connected: true }],
    activeGamepadId: 'pad-1',
  });

  assert.equal(scheduled.size, 1);

  const [hide] = scheduled.values();
  hide?.();

  assert.equal(classList.contains('hidden'), true);
  assert.equal(toast.textContent, '');
});

test('controller status indicator announces newly detected controllers after startup', () => {
  const toast = {
    textContent: '',
    classList: createClassList(['hidden']),
  };

  const indicator = createControllerStatusIndicator({ controllerStatusToast: toast } as never, {
    setTimeout: () => 1 as never,
    clearTimeout: () => {},
  });

  indicator.update({
    connectedGamepads: [{ id: 'pad-1', index: 0, mapping: 'standard', connected: true }],
    activeGamepadId: 'pad-1',
  });

  toast.classList.add('hidden');
  toast.textContent = '';

  indicator.update({
    connectedGamepads: [
      { id: 'pad-1', index: 0, mapping: 'standard', connected: true },
      { id: 'pad-2', index: 1, mapping: 'standard', connected: true },
    ],
    activeGamepadId: 'pad-1',
  });

  assert.equal(toast.classList.contains('hidden'), false);
  assert.match(toast.textContent, /pad-2/i);
});
