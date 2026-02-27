import type { BaseWindowTracker } from '../../window-trackers';

import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOverlayVisibilityRuntimeMainDepsHandler } from './overlay-visibility-runtime-main-deps';

test('overlay visibility runtime main deps builder maps state and geometry callbacks', () => {
  const calls: string[] = [];
  let trackerNotReadyWarningShown = false;
  const mainWindow = { id: 'main' } as never;
  const tracker = { id: 'tracker' } as unknown as BaseWindowTracker;

  const deps = createBuildOverlayVisibilityRuntimeMainDepsHandler({
    getMainWindow: () => mainWindow,
    getVisibleOverlayVisible: () => true,
    getWindowTracker: () => tracker,
    getTrackerNotReadyWarningShown: () => trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown) => {
      trackerNotReadyWarningShown = shown;
      calls.push(`tracker-warning:${shown}`);
    },
    updateVisibleOverlayBounds: () => calls.push('visible-bounds'),
    ensureOverlayWindowLevel: () => calls.push('ensure-level'),
    syncPrimaryOverlayWindowLayer: (layer) => calls.push(`primary-layer:${layer}`),
    enforceOverlayLayerOrder: () => calls.push('enforce-order'),
    syncOverlayShortcuts: () => calls.push('sync-shortcuts'),
    isMacOSPlatform: () => true,
    showOverlayLoadingOsd: () => calls.push('overlay-loading-osd'),
    resolveFallbackBounds: () => ({ x: 0, y: 0, width: 20, height: 20 }),
  })();

  assert.equal(deps.getMainWindow(), mainWindow);
  assert.equal(deps.getVisibleOverlayVisible(), true);
  assert.equal(deps.getTrackerNotReadyWarningShown(), false);
  deps.setTrackerNotReadyWarningShown(true);
  deps.updateVisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  deps.ensureOverlayWindowLevel(mainWindow);
  deps.syncPrimaryOverlayWindowLayer('visible');
  deps.enforceOverlayLayerOrder();
  deps.syncOverlayShortcuts();
  assert.equal(deps.isMacOSPlatform(), true);
  deps.showOverlayLoadingOsd('Overlay loading...');
  assert.deepEqual(deps.resolveFallbackBounds(), { x: 0, y: 0, width: 20, height: 20 });
  assert.equal(trackerNotReadyWarningShown, true);
  assert.deepEqual(calls, [
    'tracker-warning:true',
    'visible-bounds',
    'ensure-level',
    'primary-layer:visible',
    'enforce-order',
    'sync-shortcuts',
    'overlay-loading-osd',
  ]);
});
