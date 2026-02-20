import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createSetInvisibleOverlayVisibleHandler,
  createSetVisibleOverlayVisibleHandler,
  createToggleInvisibleOverlayHandler,
  createToggleVisibleOverlayHandler,
} from './overlay-visibility-actions';

test('set visible overlay handler forwards dependencies to core', () => {
  const calls: string[] = [];
  const setVisible = createSetVisibleOverlayVisibleHandler({
    setVisibleOverlayVisibleCore: (options) => {
      calls.push(`core:${options.visible}`);
      options.setVisibleOverlayVisibleState(options.visible);
      options.updateVisibleOverlayVisibility();
      options.updateInvisibleOverlayVisibility();
      options.syncInvisibleOverlayMousePassthrough();
      options.setMpvSubVisibility(!options.visible);
    },
    setVisibleOverlayVisibleState: (visible) => calls.push(`state:${visible}`),
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
    updateInvisibleOverlayVisibility: () => calls.push('update-invisible'),
    syncInvisibleOverlayMousePassthrough: () => calls.push('sync-mouse'),
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    isMpvConnected: () => true,
    setMpvSubVisibility: (visible) => calls.push(`mpv-sub:${visible}`),
  });

  setVisible(true);
  assert.deepEqual(calls, [
    'core:true',
    'state:true',
    'update-visible',
    'update-invisible',
    'sync-mouse',
    'mpv-sub:false',
  ]);
});

test('set invisible overlay handler forwards dependencies to core', () => {
  const calls: string[] = [];
  const setInvisible = createSetInvisibleOverlayVisibleHandler({
    setInvisibleOverlayVisibleCore: (options) => {
      calls.push(`core:${options.visible}`);
      options.setInvisibleOverlayVisibleState(options.visible);
      options.updateInvisibleOverlayVisibility();
      options.syncInvisibleOverlayMousePassthrough();
    },
    setInvisibleOverlayVisibleState: (visible) => calls.push(`state:${visible}`),
    updateInvisibleOverlayVisibility: () => calls.push('update-invisible'),
    syncInvisibleOverlayMousePassthrough: () => calls.push('sync-mouse'),
  });

  setInvisible(false);
  assert.deepEqual(calls, ['core:false', 'state:false', 'update-invisible', 'sync-mouse']);
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

test('toggle invisible overlay flips current invisible state', () => {
  const calls: string[] = [];
  let current = true;
  const toggle = createToggleInvisibleOverlayHandler({
    getInvisibleOverlayVisible: () => current,
    setInvisibleOverlayVisible: (visible) => {
      current = visible;
      calls.push(`set:${visible}`);
    },
  });

  toggle();
  toggle();
  assert.deepEqual(calls, ['set:false', 'set:true']);
});
