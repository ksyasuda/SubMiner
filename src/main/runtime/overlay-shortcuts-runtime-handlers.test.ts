import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayShortcutsRuntimeHandlers } from './overlay-shortcuts-runtime-handlers';

test('overlay shortcuts runtime handlers compose lifecycle handlers', () => {
  const calls: string[] = [];
  const runtime = createOverlayShortcutsRuntimeHandlers({
    overlayShortcutsRuntimeMainDeps: {
      overlayShortcutsRuntime: {
        registerOverlayShortcuts: () => calls.push('register'),
        unregisterOverlayShortcuts: () => calls.push('unregister'),
        syncOverlayShortcuts: () => calls.push('sync'),
        refreshOverlayShortcuts: () => calls.push('refresh'),
      },
    },
  });

  runtime.registerOverlayShortcuts();
  runtime.unregisterOverlayShortcuts();
  runtime.syncOverlayShortcuts();
  runtime.refreshOverlayShortcuts();

  assert.deepEqual(calls, ['register', 'unregister', 'sync', 'refresh']);
});
