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
    createInvisibleWindow: () => calls.push('invisible'),
    updateVisibleOverlayVisibility: () => calls.push('visible'),
    updateInvisibleOverlayVisibility: () => calls.push('invisible-visible'),
  })();

  deps.createMainWindow();
  deps.createInvisibleWindow();
  deps.updateVisibleOverlayVisibility();
  deps.updateInvisibleOverlayVisibility();
  assert.deepEqual(calls, ['main', 'invisible', 'visible', 'invisible-visible']);
});
