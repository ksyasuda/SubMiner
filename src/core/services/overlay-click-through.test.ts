import assert from 'node:assert/strict';
import test from 'node:test';

import { applyOverlayClickThrough } from './overlay-click-through';

test('applyOverlayClickThrough requests forwarding only off Windows', () => {
  const calls: Array<{ ignore: boolean; forward: boolean }> = [];
  const window = {
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      calls.push({ ignore, forward: options?.forward === true });
    },
  };

  applyOverlayClickThrough(window, true);
  applyOverlayClickThrough(window, false);

  assert.deepEqual(calls, [
    { ignore: true, forward: false },
    { ignore: true, forward: true },
  ]);
});
