import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildRestoreWindowsOnActivateMainDepsHandler,
  createBuildShouldRestoreWindowsOnActivateMainDepsHandler,
} from './app-lifecycle-main-activate';

test('should restore windows on activate deps builder maps visibility state checks', () => {
  const deps = createBuildShouldRestoreWindowsOnActivateMainDepsHandler({
    isOverlayRuntimeInitialized: () => true,
    getAllWindowCount: () => 0,
  })();

  assert.equal(deps.isOverlayRuntimeInitialized(), true);
  assert.equal(deps.getAllWindowCount(), 0);
});

test('restore windows on activate deps builder maps all restoration callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildRestoreWindowsOnActivateMainDepsHandler({
    createMainWindow: () => calls.push('main'),
    updateVisibleOverlayVisibility: () => calls.push('visible'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('mpv-sync'),
  })();

  deps.createMainWindow();
  deps.updateVisibleOverlayVisibility();
  deps.syncOverlayMpvSubtitleSuppression();
  assert.deepEqual(calls, ['main', 'visible', 'mpv-sync']);
});
