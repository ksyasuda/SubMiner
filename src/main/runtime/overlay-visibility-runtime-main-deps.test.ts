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
    getModalActive: () => true,
    getVisibleOverlayVisible: () => true,
    getForceMousePassthrough: () => true,
    getOverlayInteractionActive: () => true,
    getWindowTracker: () => tracker,
    getLastKnownWindowsForegroundProcessName: () => 'mpv',
    getWindowsOverlayProcessName: () => 'subminer',
    getWindowsFocusHandoffGraceActive: () => true,
    getTrackerNotReadyWarningShown: () => trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown) => {
      trackerNotReadyWarningShown = shown;
      calls.push(`tracker-warning:${shown}`);
    },
    updateVisibleOverlayBounds: () => calls.push('visible-bounds'),
    ensureOverlayWindowLevel: () => calls.push('ensure-level'),
    syncWindowsOverlayToMpvZOrder: () => calls.push('sync-windows-z-order'),
    syncPrimaryOverlayWindowLayer: (layer) => calls.push(`primary-layer:${layer}`),
    enforceOverlayLayerOrder: () => calls.push('enforce-order'),
    syncOverlayShortcuts: () => calls.push('sync-shortcuts'),
    isMacOSPlatform: () => true,
    isWindowsPlatform: () => false,
    showOverlayLoadingOsd: () => calls.push('overlay-loading-osd'),
    resolveFallbackBounds: () => ({ x: 0, y: 0, width: 20, height: 20 }),
  })();

  assert.equal(deps.getMainWindow(), mainWindow);
  assert.equal(deps.getModalActive(), true);
  assert.equal(deps.getVisibleOverlayVisible(), true);
  assert.equal(deps.getForceMousePassthrough(), true);
  assert.equal(deps.getOverlayInteractionActive?.(), true);
  assert.equal(deps.getLastKnownWindowsForegroundProcessName?.(), 'mpv');
  assert.equal(deps.getWindowsOverlayProcessName?.(), 'subminer');
  assert.equal(deps.getWindowsFocusHandoffGraceActive?.(), true);
  assert.equal(deps.getTrackerNotReadyWarningShown(), false);
  deps.setTrackerNotReadyWarningShown(true);
  deps.updateVisibleOverlayBounds({ x: 0, y: 0, width: 10, height: 10 });
  deps.ensureOverlayWindowLevel(mainWindow);
  deps.syncWindowsOverlayToMpvZOrder?.(mainWindow);
  deps.syncPrimaryOverlayWindowLayer('visible');
  deps.enforceOverlayLayerOrder();
  deps.syncOverlayShortcuts();
  assert.equal(deps.isMacOSPlatform(), true);
  assert.equal(deps.isWindowsPlatform(), false);
  deps.showOverlayLoadingOsd('Overlay loading...');
  assert.deepEqual(deps.resolveFallbackBounds(), { x: 0, y: 0, width: 20, height: 20 });
  assert.equal(trackerNotReadyWarningShown, true);
  assert.deepEqual(calls, [
    'tracker-warning:true',
    'visible-bounds',
    'ensure-level',
    'sync-windows-z-order',
    'primary-layer:visible',
    'enforce-order',
    'sync-shortcuts',
    'overlay-loading-osd',
  ]);
});
