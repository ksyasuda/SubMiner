import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildEnforceOverlayLayerOrderMainDepsHandler,
  createBuildEnsureOverlayWindowLevelMainDepsHandler,
  createBuildUpdateInvisibleOverlayBoundsMainDepsHandler,
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler,
} from './overlay-window-layout-main-deps';

test('overlay window layout main deps builders map callbacks', () => {
  const calls: string[] = [];

  const visible = createBuildUpdateVisibleOverlayBoundsMainDepsHandler({
    setOverlayWindowBounds: (layer) => calls.push(`visible:${layer}`),
  })();
  visible.setOverlayWindowBounds('visible', { x: 0, y: 0, width: 1, height: 1 });

  const invisible = createBuildUpdateInvisibleOverlayBoundsMainDepsHandler({
    setOverlayWindowBounds: (layer) => calls.push(`invisible:${layer}`),
  })();
  invisible.setOverlayWindowBounds('invisible', { x: 0, y: 0, width: 1, height: 1 });

  const level = createBuildEnsureOverlayWindowLevelMainDepsHandler({
    ensureOverlayWindowLevelCore: () => calls.push('ensure'),
  })();
  level.ensureOverlayWindowLevelCore({});

  const order = createBuildEnforceOverlayLayerOrderMainDepsHandler({
    enforceOverlayLayerOrderCore: () => calls.push('order'),
    getVisibleOverlayVisible: () => true,
    getInvisibleOverlayVisible: () => false,
    getMainWindow: () => ({ kind: 'main' }),
    getInvisibleWindow: () => ({ kind: 'invisible' }),
    ensureOverlayWindowLevel: () => calls.push('order-level'),
  })();
  order.enforceOverlayLayerOrderCore({
    visibleOverlayVisible: true,
    invisibleOverlayVisible: false,
    mainWindow: null,
    invisibleWindow: null,
    ensureOverlayWindowLevel: () => {},
  });
  assert.equal(order.getVisibleOverlayVisible(), true);
  assert.equal(order.getInvisibleOverlayVisible(), false);
  assert.deepEqual(order.getMainWindow(), { kind: 'main' });
  assert.deepEqual(order.getInvisibleWindow(), { kind: 'invisible' });
  order.ensureOverlayWindowLevel({});

  assert.deepEqual(calls, [
    'visible:visible',
    'invisible:invisible',
    'ensure',
    'order',
    'order-level',
  ]);
});
