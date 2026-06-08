import assert from 'node:assert/strict';
import test from 'node:test';

import { createOverlayLoadingOsdController } from './overlay-loading-osd';

test('overlay loading OSD shows spinner ticks and clears when stopped', () => {
  const messages: string[] = [];
  const clearedTimers: unknown[] = [];
  let tick: (() => void) | null = null;
  const controller = createOverlayLoadingOsdController({
    showOsd: (message) => {
      messages.push(message);
    },
    clearOsd: () => {
      messages.push('clear');
    },
    setInterval: (callback) => {
      tick = callback;
      return 'timer';
    },
    clearInterval: (timer) => {
      clearedTimers.push(timer);
    },
  });

  controller.start();
  controller.start();

  assert.deepEqual(messages, ['Overlay loading |']);
  if (!tick) {
    assert.fail('expected spinner tick callback');
  }
  const tickCallback: () => void = tick;
  tickCallback();
  tickCallback();

  controller.stop();
  controller.stop();

  assert.deepEqual(messages, ['Overlay loading |', 'Overlay loading /', 'Overlay loading -', 'clear']);
  assert.deepEqual(clearedTimers, ['timer']);
});

