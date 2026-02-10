import test from "node:test";
import assert from "node:assert/strict";
import {
  broadcastRuntimeOptionsChangedRuntimeService,
  broadcastToOverlayWindowsRuntimeService,
  getOverlayWindowsRuntimeService,
  setOverlayDebugVisualizationEnabledRuntimeService,
} from "./overlay-broadcast-runtime-service";

test("getOverlayWindowsRuntimeService returns non-destroyed windows only", () => {
  const alive = { isDestroyed: () => false } as unknown as Electron.BrowserWindow;
  const dead = { isDestroyed: () => true } as unknown as Electron.BrowserWindow;
  const windows = getOverlayWindowsRuntimeService({
    mainWindow: alive,
    invisibleWindow: dead,
  });
  assert.deepEqual(windows, [alive]);
});

test("broadcastToOverlayWindowsRuntimeService sends channel to each window", () => {
  const calls: unknown[][] = [];
  const window = {
    webContents: {
      send: (...args: unknown[]) => {
        calls.push(args);
      },
    },
  } as unknown as Electron.BrowserWindow;
  broadcastToOverlayWindowsRuntimeService([window], "x", 1, "a");
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
