import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateVisibleOverlayBoundsHandler,
} from './overlay-window-layout';

test('visible bounds handler writes visible layer geometry', () => {
  const calls: Array<{ x: number; y: number; width: number; height: number }> = [];
  const handleVisible = createUpdateVisibleOverlayBoundsHandler({
    setOverlayWindowBounds: (geometry) => calls.push(geometry),
  });
  const geometry = { x: 0, y: 0, width: 100, height: 50 };
  handleVisible(geometry);
  assert.deepEqual(calls, [geometry]);
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
      params.ensureOverlayWindowLevel({});
    },
    getVisibleOverlayVisible: () => true,
    getMainWindow: () => ({}),
    ensureOverlayWindowLevel: () => calls.push('ensure-level'),
  });
  enforce();
  assert.deepEqual(calls, ['visible-on', 'ensure-level']);
});
