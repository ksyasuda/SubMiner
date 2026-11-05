import assert from 'node:assert/strict';
import test, { type TestContext } from 'node:test';
import type { BrowserWindow, Display } from 'electron';
import type { HyprlandPlacementStatus } from '../../core/services/hyprland-window-placement';
import {
  createOverlayGeometryRuntime,
  type OverlayGeometryRuntimeDeps,
} from './overlay-geometry-runtime';

const NO_PLACEMENT: HyprlandPlacementStatus = {
  applicable: false,
  clientFound: false,
  dispatched: false,
};
const DISPLAY = { bounds: { x: 0, y: 0, width: 1920, height: 1080 } } as Display;

/** Builds the runtime with `process.platform` set and a visible overlay window that logs calls. */
function setup(t: TestContext, platform: NodeJS.Platform) {
  const originalPlatform = process.platform;
  Object.defineProperty(process, 'platform', { value: platform, configurable: true });
  t.after(() => Object.defineProperty(process, 'platform', { value: originalPlatform }));

  const windowCalls: string[] = [];
  const record =
    (name: string) =>
    (..._args: unknown[]) => {
      windowCalls.push(name);
    };
  const mainWindow = {
    isDestroyed: () => false,
    getTitle: () => 'SubMiner Overlay',
    getBounds: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    setShape: record('setShape'),
    setAlwaysOnTop: record('setAlwaysOnTop'),
    setVisibleOnAllWorkspaces: record('setVisibleOnAllWorkspaces'),
    setFullScreenable: record('setFullScreenable'),
    moveTop: record('moveTop'),
  } as unknown as BrowserWindow;
  const noop = () => {};
  const deps: OverlayGeometryRuntimeDeps = {
    screen: {
      getCursorScreenPoint: () => ({ x: 0, y: 0 }),
      getDisplayNearestPoint: () => DISPLAY,
      getDisplayMatching: () => DISPLAY,
    },
    promoteStatsOverlayAbovePlayback: noop,
    overlayManager: {
      getMainWindow: () => mainWindow,
      getModalWindow: () => null,
      getVisibleOverlayVisible: () => true,
      setOverlayWindowBounds: () => NO_PLACEMENT,
      setModalWindowBounds: () => NO_PLACEMENT,
    },
    getTrackedWindowGeometry: () => null,
    getTrackedWindowMediaSourceId: () => null,
    getTrackedWindowNativeId: () => null,
    getStatsOverlayVisible: () => false,
    getOverlayForegroundSeparateWindows: () => [],
    getLinuxVisibleOverlayWindowMode: () => 'managed',
    getLinuxTrackedMpvFullscreen: () => false,
    getLinuxTrackedMpvFullscreenChangedAtMs: () => 0,
    syncLinuxVisibleOverlayMpvFullscreenMode: noop,
    getLinuxVisibleOverlayOwnerBindingKey: () => null,
    setLinuxVisibleOverlayOwnerBindingKey: noop,
    clearVisibleOverlayX11OwnerBinding: noop,
    getNativeWindowHandleDecimal: () => '0',
    enqueueVisibleOverlayX11OwnerBindingOperation: noop,
    scheduleWindowsVisibleOverlayZOrderSyncBurst: noop,
    logDebug: noop,
  };
  return { geometry: createOverlayGeometryRuntime(deps), windowCalls };
}

// New mpv bounds can leave the X11 input shape sized for the old window; restore it to the full
// window before re-asserting the overlay's level so the whole overlay stays clickable.
test('a Linux bounds update restores the overlay window shape before raising it', (t) => {
  const { geometry, windowCalls } = setup(t, 'linux');

  geometry.updateVisibleOverlayBounds({ x: 10, y: 20, width: 1280, height: 720 });

  const shapeIndex = windowCalls.indexOf('setShape');
  assert.ok(shapeIndex >= 0, 'window shape was not restored');
  assert.ok(shapeIndex < windowCalls.indexOf('setAlwaysOnTop'));
});

test('a macOS bounds update raises the overlay without touching the window shape', (t) => {
  const { geometry, windowCalls } = setup(t, 'darwin');

  geometry.updateVisibleOverlayBounds({ x: 10, y: 20, width: 1280, height: 720 });

  assert.ok(windowCalls.includes('setAlwaysOnTop'));
  assert.ok(!windowCalls.includes('setShape'));
});
