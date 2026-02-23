import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateInvisibleOverlayBoundsHandler,
  createUpdateVisibleOverlayBoundsHandler,
} from './overlay-window-layout';

test('visible bounds handler writes visible layer geometry', () => {
  const calls: string[] = [];
  const handleVisible = createUpdateVisibleOverlayBoundsHandler({
    setOverlayWindowBounds: (layer) => calls.push(layer),
  });
  handleVisible({ x: 0, y: 0, width: 100, height: 50 });
  assert.deepEqual(calls, ['visible']);
});

test('invisible bounds handler writes invisible layer geometry', () => {
  const calls: string[] = [];
  const handleInvisible = createUpdateInvisibleOverlayBoundsHandler({
    setOverlayWindowBounds: (layer) => calls.push(layer),
  });
  handleInvisible({ x: 0, y: 0, width: 100, height: 50 });
  assert.deepEqual(calls, ['invisible']);
});

test('ensure overlay window level handler delegates to core', () => {
  const calls: string[] = [];
  const ensureLevel = createEnsureOverlayWindowLevelHandler({
    ensureOverlayWindowLevelCore: () => calls.push('core'),
  });
  ensureLevel({});
  assert.deepEqual(calls, ['core']);
});

test('enforce overlay layer order handler forwards resolved state', () => {
  const calls: string[] = [];
  const enforce = createEnforceOverlayLayerOrderHandler({
    enforceOverlayLayerOrderCore: (params) => {
      calls.push(params.visibleOverlayVisible ? 'visible-on' : 'visible-off');
      calls.push(params.invisibleOverlayVisible ? 'invisible-on' : 'invisible-off');
      params.ensureOverlayWindowLevel({});
    },
    getVisibleOverlayVisible: () => true,
    getInvisibleOverlayVisible: () => false,
    getMainWindow: () => ({}),
    getInvisibleWindow: () => ({}),
    ensureOverlayWindowLevel: () => calls.push('ensure-level'),
  });
  enforce();
  assert.deepEqual(calls, ['visible-on', 'invisible-off', 'ensure-level']);
});
