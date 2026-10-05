import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildEnsureTrayMainDepsHandler } from './app-runtime-main-deps';

test('ensure tray main deps trigger overlay bootstrap on tray click when runtime not initialized', () => {
  const calls: string[] = [];
  const deps = createBuildEnsureTrayMainDepsHandler({
    getTray: () => null,
    setTray: () => calls.push('set-tray'),
    buildTrayMenu: () => ({}),
    resolveTrayIconPath: () => null,
    createImageFromPath: () => ({}),
    createEmptyImage: () => ({}),
    createTray: () => ({}),
    trayTooltip: 'SubMiner',
    platform: 'darwin',
    logWarn: (message) => calls.push(`warn:${message}`),
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    isOverlayRuntimeInitialized: () => false,
    setVisibleOverlayVisible: (visible) => calls.push(`set-visible:${visible}`),
  })();

  deps.ensureOverlayVisibleFromTrayClick();
  assert.deepEqual(calls, ['init-overlay', 'set-visible:true']);
});
