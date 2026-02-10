import test from "node:test";
import assert from "node:assert/strict";
import { createOverlayManagerService } from "./overlay-manager-service";

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
