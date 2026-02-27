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
    setOverlayWindowBounds: () => calls.push('visible'),
  })();
  visible.setOverlayWindowBounds({ x: 0, y: 0, width: 1, height: 1 });

  const level = createBuildEnsureOverlayWindowLevelMainDepsHandler({
    ensureOverlayWindowLevelCore: () => calls.push('ensure'),
  })();
  level.ensureOverlayWindowLevelCore({});

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
    'visible',
    'ensure',
    'order',
    'order-level',
  ]);
});
