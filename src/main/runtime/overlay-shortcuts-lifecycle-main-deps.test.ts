import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildRefreshOverlayShortcutsMainDepsHandler,
  createBuildRegisterOverlayShortcutsMainDepsHandler,
  createBuildSyncOverlayShortcutsMainDepsHandler,
  createBuildUnregisterOverlayShortcutsMainDepsHandler,
} from './overlay-shortcuts-lifecycle-main-deps';

test('overlay shortcuts lifecycle main deps builders map runtime instance', () => {
  const runtime = {
    registerOverlayShortcuts: () => {},
    unregisterOverlayShortcuts: () => {},
    syncOverlayShortcuts: () => {},
    refreshOverlayShortcuts: () => {},
  };

  const register = createBuildRegisterOverlayShortcutsMainDepsHandler({
    overlayShortcutsRuntime: runtime,
  })();
  const unregister = createBuildUnregisterOverlayShortcutsMainDepsHandler({
    overlayShortcutsRuntime: runtime,
  })();
  const sync = createBuildSyncOverlayShortcutsMainDepsHandler({
    overlayShortcutsRuntime: runtime,
  })();
  const refresh = createBuildRefreshOverlayShortcutsMainDepsHandler({
    overlayShortcutsRuntime: runtime,
  })();

  assert.equal(register.overlayShortcutsRuntime, runtime);
  assert.equal(unregister.overlayShortcutsRuntime, runtime);
  assert.equal(sync.overlayShortcutsRuntime, runtime);
  assert.equal(refresh.overlayShortcutsRuntime, runtime);
});
