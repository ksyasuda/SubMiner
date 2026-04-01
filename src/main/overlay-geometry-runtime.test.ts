import assert from 'node:assert/strict';
import test from 'node:test';

import { createOverlayGeometryRuntime } from './overlay-geometry-runtime';

test('overlay geometry runtime prefers tracker geometry before fallback', () => {
  const overlayBounds: unknown[] = [];
  const modalBounds: unknown[] = [];
  const layerCalls: Array<[unknown, unknown]> = [];
  const levelCalls: unknown[] = [];

  const runtime = createOverlayGeometryRuntime({
    screen: {
      getCursorScreenPoint: () => ({ x: 1, y: 2 }),
      getDisplayNearestPoint: () => ({
        workArea: { x: 10, y: 20, width: 30, height: 40 },
      }),
    },
    windowState: {
      getMainWindow: () =>
        ({
          isDestroyed: () => false,
        }) as never,
      setOverlayWindowBounds: (geometry) => overlayBounds.push(geometry),
      setModalWindowBounds: (geometry) => modalBounds.push(geometry),
      getVisibleOverlayVisible: () => true,
    },
    getWindowTracker: () => ({
      getGeometry: () => ({ x: 100, y: 200, width: 300, height: 400 }),
    }),
    ensureOverlayWindowLevelCore: (window) => {
      levelCalls.push(window);
    },
    syncOverlayWindowLayer: (window, layer) => {
      layerCalls.push([window, layer]);
    },
    enforceOverlayLayerOrderCore: ({
      visibleOverlayVisible,
      mainWindow,
      ensureOverlayWindowLevel,
    }) => {
      if (visibleOverlayVisible && mainWindow) {
        ensureOverlayWindowLevel(mainWindow);
      }
    },
  });

  assert.deepEqual(runtime.getCurrentOverlayGeometry(), {
    x: 100,
    y: 200,
    width: 300,
    height: 400,
  });
  assert.equal(
    runtime.geometryMatches(
      { x: 1, y: 2, width: 3, height: 4 },
      { x: 1, y: 2, width: 3, height: 4 },
    ),
    true,
  );
  assert.equal(runtime.geometryMatches({ x: 1, y: 2, width: 3, height: 4 }, null), false);

  runtime.applyOverlayRegions({ x: 7, y: 8, width: 9, height: 10 });
  assert.deepEqual(overlayBounds, [{ x: 7, y: 8, width: 9, height: 10 }]);
  assert.deepEqual(modalBounds, [{ x: 7, y: 8, width: 9, height: 10 }]);

  runtime.syncPrimaryOverlayWindowLayer('visible');
  runtime.ensureOverlayWindowLevel({
    isDestroyed: () => false,
  } as never);
  runtime.enforceOverlayLayerOrder();

  assert.equal(layerCalls.length >= 1, true);
  assert.equal(levelCalls.length >= 2, true);
});
