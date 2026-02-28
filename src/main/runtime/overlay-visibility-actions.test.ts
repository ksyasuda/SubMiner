import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createSetVisibleOverlayVisibleHandler,
  createToggleVisibleOverlayHandler,
} from './overlay-visibility-actions';

test('set visible overlay handler forwards dependencies to core', () => {
  const calls: string[] = [];
  let warmupStarts = 0;
  const setVisible = createSetVisibleOverlayVisibleHandler({
    setVisibleOverlayVisibleCore: (options) => {
      calls.push(`core:${options.visible}`);
      options.setVisibleOverlayVisibleState(options.visible);
      options.updateVisibleOverlayVisibility();
    },
    setVisibleOverlayVisibleState: (visible) => calls.push(`state:${visible}`),
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
    onVisibleOverlayEnabled: () => {
      warmupStarts += 1;
    },
  });

  setVisible(true);
  assert.deepEqual(calls, [
    'core:true',
    'state:true',
    'update-visible',
  ]);
  assert.equal(warmupStarts, 1);

  setVisible(false);
  assert.equal(warmupStarts, 1);
});

test('toggle visible overlay flips current visible state', () => {
  const calls: string[] = [];
  let current = false;
  const toggle = createToggleVisibleOverlayHandler({
    getVisibleOverlayVisible: () => current,
    setVisibleOverlayVisible: (visible) => {
      current = visible;
      calls.push(`set:${visible}`);
    },
  });

  toggle();
  toggle();
  assert.deepEqual(calls, ['set:true', 'set:false']);
});
