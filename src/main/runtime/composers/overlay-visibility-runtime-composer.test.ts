import assert from 'node:assert/strict';
import test from 'node:test';
import { composeOverlayVisibilityRuntime } from './overlay-visibility-runtime-composer';

test('composeOverlayVisibilityRuntime returns overlay visibility handlers', () => {
  const calls: string[] = [];
  const composed = composeOverlayVisibilityRuntime({
    overlayVisibilityRuntime: {
      updateVisibleOverlayVisibility: () => {
        calls.push('updateVisibleOverlayVisibility');
      },
    },
    restorePreviousSecondarySubVisibilityMainDeps: {
      getMpvClient: () => null,
    },
    broadcastRuntimeOptionsChangedMainDeps: {
      broadcastRuntimeOptionsChangedRuntime: () => {
        calls.push('broadcastRuntimeOptionsChangedRuntime');
      },
      getRuntimeOptionsState: () => [],
      broadcastToOverlayWindows: () => {},
    },
    sendToActiveOverlayWindowMainDeps: {
      sendToActiveOverlayWindowRuntime: () => true,
    },
    setOverlayDebugVisualizationEnabledMainDeps: {
      setOverlayDebugVisualizationEnabledRuntime: () => {},
      getCurrentEnabled: () => false,
      setCurrentEnabled: () => {},
    },
    openRuntimeOptionsPaletteMainDeps: {
      openRuntimeOptionsPaletteRuntime: () => {
        calls.push('openRuntimeOptionsPaletteRuntime');
      },
    },
  });

  assert.equal(typeof composed.updateVisibleOverlayVisibility, 'function');
  assert.equal(typeof composed.restorePreviousSecondarySubVisibility, 'function');
  assert.equal(typeof composed.broadcastRuntimeOptionsChanged, 'function');
  assert.equal(typeof composed.sendToActiveOverlayWindow, 'function');
  assert.equal(typeof composed.setOverlayDebugVisualizationEnabled, 'function');
  assert.equal(typeof composed.openRuntimeOptionsPalette, 'function');

  // updateVisibleOverlayVisibility passes through to the injected runtime dep
  composed.updateVisibleOverlayVisibility();
  assert.deepEqual(calls, ['updateVisibleOverlayVisibility']);

  // openRuntimeOptionsPalette forwards to the injected runtime dep
  composed.openRuntimeOptionsPalette();
  assert.deepEqual(calls, ['updateVisibleOverlayVisibility', 'openRuntimeOptionsPaletteRuntime']);

  // broadcastRuntimeOptionsChanged forwards to the injected runtime dep
  composed.broadcastRuntimeOptionsChanged();
  assert.ok(calls.includes('broadcastRuntimeOptionsChangedRuntime'));
});
