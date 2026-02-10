import test from "node:test";
import assert from "node:assert/strict";
import {
  broadcastRuntimeOptionsChangedRuntimeService,
  createOverlayManagerService,
  setOverlayDebugVisualizationEnabledRuntimeService,
} from "./overlay-manager-service";

test("overlay manager initializes with empty windows and hidden overlays", () => {
  const manager = createOverlayManagerService();
  assert.equal(manager.getMainWindow(), null);
  assert.equal(manager.getInvisibleWindow(), null);
  assert.equal(manager.getVisibleOverlayVisible(), false);
  assert.equal(manager.getInvisibleOverlayVisible(), false);
  assert.deepEqual(manager.getOverlayWindows(), []);
});

test("overlay manager stores window references and returns stable window order", () => {
  const manager = createOverlayManagerService();
  const visibleWindow = { isDestroyed: () => false } as unknown as Electron.BrowserWindow;
  const invisibleWindow = { isDestroyed: () => false } as unknown as Electron.BrowserWindow;

  manager.setMainWindow(visibleWindow);
  manager.setInvisibleWindow(invisibleWindow);

  assert.equal(manager.getMainWindow(), visibleWindow);
  assert.equal(manager.getInvisibleWindow(), invisibleWindow);
  assert.deepEqual(manager.getOverlayWindows(), [visibleWindow, invisibleWindow]);
});

test("overlay manager excludes destroyed windows", () => {
  const manager = createOverlayManagerService();
  manager.setMainWindow({ isDestroyed: () => true } as unknown as Electron.BrowserWindow);
  manager.setInvisibleWindow({ isDestroyed: () => false } as unknown as Electron.BrowserWindow);

  assert.equal(manager.getOverlayWindows().length, 1);
});

test("overlay manager stores visibility state", () => {
  const manager = createOverlayManagerService();

  manager.setVisibleOverlayVisible(true);
  manager.setInvisibleOverlayVisible(true);
  assert.equal(manager.getVisibleOverlayVisible(), true);
  assert.equal(manager.getInvisibleOverlayVisible(), true);
});

test("overlay manager broadcasts to non-destroyed windows", () => {
  const manager = createOverlayManagerService();
  const calls: unknown[][] = [];
  const aliveWindow = {
    isDestroyed: () => false,
    webContents: {
      send: (...args: unknown[]) => {
        calls.push(args);
      },
    },
  } as unknown as Electron.BrowserWindow;
  const deadWindow = {
    isDestroyed: () => true,
    webContents: {
      send: (..._args: unknown[]) => {},
    },
  } as unknown as Electron.BrowserWindow;

  manager.setMainWindow(aliveWindow);
  manager.setInvisibleWindow(deadWindow);
  manager.broadcastToOverlayWindows("x", 1, "a");

  assert.deepEqual(calls, [["x", 1, "a"]]);
});

test("runtime-option and debug broadcasts use expected channels", () => {
  const broadcasts: unknown[][] = [];
  broadcastRuntimeOptionsChangedRuntimeService(
    () => [],
    (channel, ...args) => {
      broadcasts.push([channel, ...args]);
    },
  );
  let state = false;
  const changed = setOverlayDebugVisualizationEnabledRuntimeService(
    state,
    true,
    (enabled) => {
      state = enabled;
    },
    (channel, ...args) => {
      broadcasts.push([channel, ...args]);
    },
  );
  assert.equal(changed, true);
  assert.equal(state, true);
  assert.deepEqual(broadcasts, [
    ["runtime-options:changed", []],
    ["overlay-debug-visualization:set", true],
  ]);
});
