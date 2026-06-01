import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateVisibleOverlayBoundsHandler,
  hasLiveOverlayWindowBoundsMismatch,
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

test('visible bounds handler runs follow-up callback after applying geometry', () => {
  const calls: string[] = [];
  const geometry = { x: 0, y: 0, width: 100, height: 50 };
  const handleVisible = createUpdateVisibleOverlayBoundsHandler({
    setOverlayWindowBounds: () => calls.push('set-bounds'),
    afterSetOverlayWindowBounds: (nextGeometry) => {
      assert.deepEqual(nextGeometry, geometry);
      calls.push('after-bounds');
    },
  });

  handleVisible(geometry);

  assert.deepEqual(calls, ['set-bounds', 'after-bounds']);
});

test('visible bounds handler skips unchanged geometry', () => {
  const calls: string[] = [];
  const geometry = { x: 0, y: 0, width: 100, height: 50 };
  const handleVisible = createUpdateVisibleOverlayBoundsHandler({
    getCurrentOverlayWindowBounds: () => ({ ...geometry }),
    setOverlayWindowBounds: () => calls.push('set-bounds'),
    afterSetOverlayWindowBounds: () => calls.push('after-bounds'),
  });

  handleVisible(geometry);

  assert.deepEqual(calls, []);
});

test('visible bounds handler can refresh unchanged geometry for mode reconciliation', () => {
  const calls: string[] = [];
  const geometry = { x: 0, y: 0, width: 100, height: 50 };
  const handleVisible = createUpdateVisibleOverlayBoundsHandler({
    getCurrentOverlayWindowBounds: () => ({ ...geometry }),
    shouldRefreshUnchangedGeometry: (nextGeometry) => {
      assert.deepEqual(nextGeometry, geometry);
      calls.push('refresh-check');
      return true;
    },
    setOverlayWindowBounds: () => calls.push('set-bounds'),
    afterSetOverlayWindowBounds: () => calls.push('after-bounds'),
  });

  handleVisible(geometry);

  assert.deepEqual(calls, ['refresh-check', 'set-bounds', 'after-bounds']);
});

test('live overlay bounds mismatch forces refresh after window manager restore drift', () => {
  const geometry = { x: 100, y: 80, width: 1280, height: 720 };

  assert.equal(
    hasLiveOverlayWindowBoundsMismatch(
      [
        {
          isDestroyed: () => false,
          getBounds: () => ({ x: 96, y: 76, width: 1300, height: 740 }),
        },
      ],
      geometry,
    ),
    true,
  );
  assert.equal(
    hasLiveOverlayWindowBoundsMismatch(
      [
        {
          isDestroyed: () => false,
          getBounds: () => ({ ...geometry }),
        },
        {
          isDestroyed: () => true,
          getBounds: () => ({ x: 0, y: 0, width: 1, height: 1 }),
        },
      ],
      geometry,
    ),
    false,
  );
});

test('ensure overlay window level handler delegates to core', () => {
  const calls: string[] = [];
  const ensureLevel = createEnsureOverlayWindowLevelHandler({
    ensureOverlayWindowLevelCore: () => calls.push('core'),
    afterEnsureOverlayWindowLevel: () => calls.push('after'),
  });
  ensureLevel({});
  assert.deepEqual(calls, ['core', 'after']);
});

test('ensure overlay window level handler skips while top reassertion is suppressed', () => {
  const calls: string[] = [];
  const window = {};
  const ensureLevel = createEnsureOverlayWindowLevelHandler({
    shouldSuppressOverlayWindowLevel: (nextWindow) => {
      assert.equal(nextWindow, window);
      calls.push('suppress-check');
      return true;
    },
    ensureOverlayWindowLevelCore: () => calls.push('core'),
    afterEnsureOverlayWindowLevel: () => calls.push('after'),
  });

  ensureLevel(window);

  assert.deepEqual(calls, ['suppress-check']);
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
