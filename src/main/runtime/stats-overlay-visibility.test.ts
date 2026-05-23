import assert from 'node:assert/strict';
import test from 'node:test';

import { createStatsOverlayVisibilityChangeHandler } from './stats-overlay-visibility';

test('stats overlay visibility handler makes overlay mouse-passive before opening stats', () => {
  const calls: string[] = [];
  const handler = createStatsOverlayVisibilityChangeHandler({
    setStatsOverlayVisibleState: (visible) => calls.push(`state:${visible}`),
    resetVisibleOverlayInteraction: () => calls.push('reset-interaction'),
    getMainWindow: () =>
      ({
        isDestroyed: () => false,
        isVisible: () => true,
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          calls.push(`mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`);
        },
      }) as never,
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
  });

  handler(true);

  assert.deepEqual(calls, [
    'state:true',
    'reset-interaction',
    'mouse-ignore:true:forward',
    'update-visible',
  ]);
});

test('stats overlay visibility handler restores overlay then leaves mpv mouse-responsive after close', () => {
  const calls: string[] = [];
  const handler = createStatsOverlayVisibilityChangeHandler({
    setStatsOverlayVisibleState: (visible) => calls.push(`state:${visible}`),
    resetVisibleOverlayInteraction: () => calls.push('reset-interaction'),
    getMainWindow: () =>
      ({
        isDestroyed: () => false,
        isVisible: () => true,
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          calls.push(`mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`);
        },
      }) as never,
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
  });

  handler(false);

  assert.deepEqual(calls, [
    'state:false',
    'reset-interaction',
    'update-visible',
    'mouse-ignore:true:forward',
  ]);
});
