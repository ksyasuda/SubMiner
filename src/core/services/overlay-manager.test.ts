import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createOverlayManager,
  setOverlayDebugVisualizationEnabledRuntime,
} from './overlay-manager';

test('overlay manager excludes destroyed windows', () => {
  const manager = createOverlayManager();
  manager.setMainWindow({
    isDestroyed: () => true,
  } as unknown as Electron.BrowserWindow);
  manager.setModalWindow({
    isDestroyed: () => false,
  } as unknown as Electron.BrowserWindow);

  assert.equal(manager.getOverlayWindows().length, 0);
});

test('overlay manager broadcasts to non-destroyed windows', () => {
  const manager = createOverlayManager();
  const calls: unknown[][] = [];
  const aliveWindow = {
    isDestroyed: () => false,
    webContents: {
      send: (...args: unknown[]) => {
        calls.push(args);
      },
    },
  } as unknown as Electron.BrowserWindow;
  manager.setMainWindow(aliveWindow);
  manager.setModalWindow({
    isDestroyed: () => false,
    webContents: { send: () => {} },
  } as unknown as Electron.BrowserWindow);
  manager.broadcastToOverlayWindows('x', 1, 'a');

  assert.deepEqual(calls, [['x', 1, 'a']]);
});

test('overlay manager applies bounds for main and modal windows', () => {
  const manager = createOverlayManager();
  const visibleCalls: Electron.Rectangle[] = [];
  const visibleWindow = {
    isDestroyed: () => false,
    getTitle: () => 'SubMiner Overlay',
    setBounds: (bounds: Electron.Rectangle) => {
      visibleCalls.push(bounds);
    },
  } as unknown as Electron.BrowserWindow;
  const modalCalls: Electron.Rectangle[] = [];
  const modalWindow = {
    isDestroyed: () => false,
    getTitle: () => 'SubMiner Overlay Modal',
    setBounds: (bounds: Electron.Rectangle) => {
      modalCalls.push(bounds);
    },
  } as unknown as Electron.BrowserWindow;
  manager.setMainWindow(visibleWindow);
  manager.setModalWindow(modalWindow);

  manager.setOverlayWindowBounds({
    x: 10,
    y: 20,
    width: 30,
    height: 40,
  });
  manager.setModalWindowBounds({
    x: 80,
    y: 90,
    width: 100,
    height: 110,
  });

  assert.deepEqual(visibleCalls, [{ x: 10, y: 20, width: 30, height: 40 }]);
  assert.deepEqual(modalCalls, [{ x: 80, y: 90, width: 100, height: 110 }]);
});

test('overlay manager can suppress z-order promotion during bounds updates', () => {
  const calls: string[] = [];
  const createManager = createOverlayManager as unknown as (options: {
    updateOverlayWindowBounds: (
      geometry: Electron.Rectangle,
      window: Electron.BrowserWindow | null,
      options: { promote: boolean },
    ) => void;
    shouldPromoteWindowOnBoundsUpdate: (window: Electron.BrowserWindow) => boolean;
  }) => ReturnType<typeof createOverlayManager>;
  const manager = createManager({
    updateOverlayWindowBounds: (_geometry, _window, options) => {
      calls.push(`promote:${options.promote}`);
    },
    shouldPromoteWindowOnBoundsUpdate: () => false,
  });

  manager.setMainWindow({
    isDestroyed: () => false,
  } as unknown as Electron.BrowserWindow);

  manager.setOverlayWindowBounds({ x: 1, y: 2, width: 3, height: 4 });

  assert.deepEqual(calls, ['promote:false']);
});

test('setOverlayDebugVisualizationEnabledRuntime only updates state when the value changes', () => {
  let state = false;
  const setState = (enabled: boolean) => {
    state = enabled;
  };

  assert.equal(setOverlayDebugVisualizationEnabledRuntime(state, false, setState), false);
  assert.equal(setOverlayDebugVisualizationEnabledRuntime(state, true, setState), true);
  assert.equal(state, true);
});
