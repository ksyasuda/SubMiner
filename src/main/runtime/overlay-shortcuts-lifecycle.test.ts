import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createRefreshOverlayShortcutsHandler,
  createRegisterOverlayShortcutsHandler,
  createSyncOverlayShortcutsHandler,
  createUnregisterOverlayShortcutsHandler,
} from './overlay-shortcuts-lifecycle';

function createRuntime(calls: string[]) {
  return {
    registerOverlayShortcuts: () => calls.push('register'),
    unregisterOverlayShortcuts: () => calls.push('unregister'),
    syncOverlayShortcuts: () => calls.push('sync'),
    refreshOverlayShortcuts: () => calls.push('refresh'),
  };
}

test('register overlay shortcuts handler delegates to runtime', () => {
  const calls: string[] = [];
  createRegisterOverlayShortcutsHandler({
    overlayShortcutsRuntime: createRuntime(calls),
  })();
  assert.deepEqual(calls, ['register']);
});

test('unregister overlay shortcuts handler delegates to runtime', () => {
  const calls: string[] = [];
  createUnregisterOverlayShortcutsHandler({
    overlayShortcutsRuntime: createRuntime(calls),
  })();
  assert.deepEqual(calls, ['unregister']);
});

test('sync overlay shortcuts handler delegates to runtime', () => {
  const calls: string[] = [];
  createSyncOverlayShortcutsHandler({
    overlayShortcutsRuntime: createRuntime(calls),
  })();
  assert.deepEqual(calls, ['sync']);
});

test('refresh overlay shortcuts handler delegates to runtime', () => {
  const calls: string[] = [];
  createRefreshOverlayShortcutsHandler({
    overlayShortcutsRuntime: createRuntime(calls),
  })();
  assert.deepEqual(calls, ['refresh']);
});
