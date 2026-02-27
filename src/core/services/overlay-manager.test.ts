import test from 'node:test';
import assert from 'node:assert/strict';
import {
  broadcastRuntimeOptionsChangedRuntime,
  createOverlayManager,
  setOverlayDebugVisualizationEnabledRuntime,
} from './overlay-manager';

test('overlay manager initializes with empty windows and hidden overlays', () => {
  const manager = createOverlayManager();
  assert.equal(manager.getMainWindow(), null);
  assert.equal(manager.getModalWindow(), null);
  assert.equal(manager.getVisibleOverlayVisible(), false);
  assert.deepEqual(manager.getOverlayWindows(), []);
});

test('overlay manager stores window references and returns stable window order', () => {
  const manager = createOverlayManager();
  const visibleWindow = {
    isDestroyed: () => false,
  } as unknown as Electron.BrowserWindow;
  const modalWindow = {
    isDestroyed: () => false,
  } as unknown as Electron.BrowserWindow;

  manager.setMainWindow(visibleWindow);
  manager.setModalWindow(modalWindow);

  assert.equal(manager.getMainWindow(), visibleWindow);
  assert.equal(manager.getModalWindow(), modalWindow);
  assert.equal(manager.getOverlayWindow(), visibleWindow);
  assert.deepEqual(manager.getOverlayWindows(), [visibleWindow]);
});

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

test('overlay manager stores visibility state', () => {
  const manager = createOverlayManager();

  manager.setVisibleOverlayVisible(true);
  assert.equal(manager.getVisibleOverlayVisible(), true);
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
    setBounds: (bounds: Electron.Rectangle) => {
      visibleCalls.push(bounds);
    },
  } as unknown as Electron.BrowserWindow;
  const modalCalls: Electron.Rectangle[] = [];
  const modalWindow = {
    isDestroyed: () => false,
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

test('runtime-option broadcast still uses expected channel', () => {
  const broadcasts: unknown[][] = [];
  broadcastRuntimeOptionsChangedRuntime(
    () => [],
    (channel, ...args) => {
      broadcasts.push([channel, ...args]);
    },
  );
  let state = false;
  const changed = setOverlayDebugVisualizationEnabledRuntime(
    state,
    true,
    (enabled) => {
      state = enabled;
    },
  );
  assert.equal(changed, true);
  assert.equal(state, true);
  assert.deepEqual(broadcasts, [['runtime-options:changed', []]]);
});
