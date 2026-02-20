import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOverlayVisibilityRuntimeMainDepsHandler } from './overlay-visibility-runtime-main-deps';

test('overlay visibility runtime main deps builder maps state and geometry callbacks', () => {
  const calls: string[] = [];
  let trackerNotReadyWarningShown = false;
  const mainWindow = { id: 'main' } as never;
  const invisibleWindow = { id: 'invisible' } as never;

  const deps = createBuildOverlayVisibilityRuntimeMainDepsHandler({
    getMainWindow: () => mainWindow,
    getInvisibleWindow: () => invisibleWindow,
    getVisibleOverlayVisible: () => true,
    getInvisibleOverlayVisible: () => false,
    getWindowTracker: () => ({ id: 'tracker' }),
    getTrackerNotReadyWarningShown: () => trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown) => {
      trackerNotReadyWarningShown = shown;
      calls.push(`tracker-warning:${shown}`);
    },
    updateVisibleOverlayBounds: () => calls.push('visible-bounds'),
    updateInvisibleOverlayBounds: () => calls.push('invisible-bounds'),
    ensureOverlayWindowLevel: () => calls.push('ensure-level'),
    enforceOverlayLayerOrder: () => calls.push('enforce-order'),
    syncOverlayShortcuts: () => calls.push('sync-shortcuts'),
  })();

  assert.equal(deps.getMainWindow(), mainWindow);
  assert.equal(deps.getInvisibleWindow(), invisibleWindow);
  assert.equal(deps.getVisibleOverlayVisible(), true);
  assert.equal(deps.getInvisibleOverlayVisible(), false);
  assert.equal(deps.getTrackerNotReadyWarningShown(), false);
  deps.setTrackerNotReadyWarningShown(true);
  deps.updateVisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  deps.updateInvisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  deps.ensureOverlayWindowLevel(mainWindow);
  deps.enforceOverlayLayerOrder();
  deps.syncOverlayShortcuts();
  assert.equal(trackerNotReadyWarningShown, true);
  assert.deepEqual(calls, [
    'tracker-warning:true',
    'visible-bounds',
    'invisible-bounds',
    'ensure-level',
    'enforce-order',
    'sync-shortcuts',
  ]);
});
