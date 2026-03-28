import assert from 'node:assert/strict';
import test from 'node:test';
import { composeOverlayVisibilityRuntime } from './overlay-visibility-runtime-composer';

test('composeOverlayVisibilityRuntime returns overlay visibility handlers', () => {
  const composed = composeOverlayVisibilityRuntime({
    overlayVisibilityRuntime: {
      updateVisibleOverlayVisibility: () => {},
    },
    restorePreviousSecondarySubVisibilityMainDeps: {
      getMpvClient: () => null,
    },
    broadcastRuntimeOptionsChangedMainDeps: {
      broadcastRuntimeOptionsChangedRuntime: () => {},
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
      openRuntimeOptionsPaletteRuntime: () => {},
    },
  });

  assert.equal(typeof composed.updateVisibleOverlayVisibility, 'function');
  assert.equal(typeof composed.restorePreviousSecondarySubVisibility, 'function');
  assert.equal(typeof composed.broadcastRuntimeOptionsChanged, 'function');
  assert.equal(typeof composed.sendToActiveOverlayWindow, 'function');
  assert.equal(typeof composed.setOverlayDebugVisualizationEnabled, 'function');
  assert.equal(typeof composed.openRuntimeOptionsPalette, 'function');
});
