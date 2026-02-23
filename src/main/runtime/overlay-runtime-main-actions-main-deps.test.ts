import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler,
  createBuildGetRuntimeOptionsStateMainDepsHandler,
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler,
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler,
  createBuildSendToActiveOverlayWindowMainDepsHandler,
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler,
} from './overlay-runtime-main-actions-main-deps';

test('get runtime options state main deps builder maps callbacks', () => {
  const manager = { listOptions: () => [] };
  const deps = createBuildGetRuntimeOptionsStateMainDepsHandler({
    getRuntimeOptionsManager: () => manager,
  })();
  assert.equal(deps.getRuntimeOptionsManager(), manager);
});

test('restore secondary sub visibility main deps builder maps callbacks', () => {
  const deps = createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler({
    getMpvClient: () => ({ connected: true, restorePreviousSecondarySubVisibility: () => {} }),
  })();
  assert.equal(deps.getMpvClient()?.connected, true);
});

test('broadcast runtime options changed main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildBroadcastRuntimeOptionsChangedMainDepsHandler({
    broadcastRuntimeOptionsChangedRuntime: () => calls.push('broadcast-runtime'),
    getRuntimeOptionsState: () => [],
    broadcastToOverlayWindows: (channel) => calls.push(channel),
  })();

  deps.broadcastRuntimeOptionsChangedRuntime(() => [], () => {});
  deps.broadcastToOverlayWindows('runtime-options:changed');
  assert.deepEqual(deps.getRuntimeOptionsState(), []);
  assert.deepEqual(calls, ['broadcast-runtime', 'runtime-options:changed']);
});

test('send to active overlay window main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildSendToActiveOverlayWindowMainDepsHandler({
    sendToActiveOverlayWindowRuntime: () => {
      calls.push('send');
      return true;
    },
  })();

  assert.equal(deps.sendToActiveOverlayWindowRuntime('x'), true);
  assert.deepEqual(calls, ['send']);
});

test('set overlay debug visualization main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler({
    setOverlayDebugVisualizationEnabledRuntime: () => calls.push('set-runtime'),
    getCurrentEnabled: () => false,
    setCurrentEnabled: () => calls.push('set-current'),
    broadcastToOverlayWindows: () => calls.push('broadcast'),
  })();

  deps.setOverlayDebugVisualizationEnabledRuntime(false, true, () => {}, () => {});
  assert.equal(deps.getCurrentEnabled(), false);
  deps.setCurrentEnabled(true);
  deps.broadcastToOverlayWindows('overlay:debug');
  assert.deepEqual(calls, ['set-runtime', 'set-current', 'broadcast']);
});

test('open runtime options palette main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildOpenRuntimeOptionsPaletteMainDepsHandler({
    openRuntimeOptionsPaletteRuntime: () => calls.push('open'),
  })();

  deps.openRuntimeOptionsPaletteRuntime();
  assert.deepEqual(calls, ['open']);
});
