import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildEnforceOverlayLayerOrderMainDepsHandler,
  createBuildEnsureOverlayWindowLevelMainDepsHandler,
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler,
} from './overlay-window-layout-main-deps';

test('overlay window layout main deps builders map callbacks', () => {
  const calls: string[] = [];

  const visible = createBuildUpdateVisibleOverlayBoundsMainDepsHandler({
    getCurrentOverlayWindowBounds: () => {
      calls.push('visible-current');
      return null;
    },
    setOverlayWindowBounds: () => calls.push('visible'),
  })();
  assert.equal(visible.getCurrentOverlayWindowBounds?.(), null);
  visible.setOverlayWindowBounds({ x: 0, y: 0, width: 1, height: 1 });

  const level = createBuildEnsureOverlayWindowLevelMainDepsHandler({
    shouldSuppressOverlayWindowLevel: () => {
      calls.push('ensure-suppressed-check');
      return false;
    },
    ensureOverlayWindowLevelCore: () => calls.push('ensure'),
    afterEnsureOverlayWindowLevel: () => calls.push('ensure-after'),
  })();
  assert.equal(level.shouldSuppressOverlayWindowLevel?.({}), false);
  level.ensureOverlayWindowLevelCore({});
  level.afterEnsureOverlayWindowLevel?.({});

  const order = createBuildEnforceOverlayLayerOrderMainDepsHandler({
    enforceOverlayLayerOrderCore: () => calls.push('order'),
    getVisibleOverlayVisible: () => true,
    getMainWindow: () => ({ kind: 'main' }),
    ensureOverlayWindowLevel: () => calls.push('order-level'),
  })();
  order.enforceOverlayLayerOrderCore({
    visibleOverlayVisible: true,
    mainWindow: null,
    ensureOverlayWindowLevel: () => {},
  });
  assert.equal(order.getVisibleOverlayVisible(), true);
  assert.deepEqual(order.getMainWindow(), { kind: 'main' });
  order.ensureOverlayWindowLevel({});

  assert.deepEqual(calls, [
    'visible-current',
    'visible',
    'ensure-suppressed-check',
    'ensure',
    'ensure-after',
    'order',
    'order-level',
  ]);
});
