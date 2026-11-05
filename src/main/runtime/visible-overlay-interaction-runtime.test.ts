import assert from 'node:assert/strict';
import test, { type TestContext } from 'node:test';
import type { BrowserWindow } from 'electron';
import { stopLinuxOverlayPointerInteractionLoop } from './linux-overlay-pointer-interaction';
import { stopLinuxOverlayZOrderKeepAliveLoop } from './linux-overlay-zorder-keepalive';
import {
  createVisibleOverlayInteractionRuntime,
  type VisibleOverlayInteractionRuntimeDeps,
} from './visible-overlay-interaction-runtime';

type MouseEventsCall = { ignore: boolean; options?: { forward?: boolean } };

/** Builds the runtime with `process.platform` set, a visible overlay and a fake main window. */
function setup(t: TestContext, platform: NodeJS.Platform) {
  const originalPlatform = process.platform;
  Object.defineProperty(process, 'platform', { value: platform, configurable: true });
  t.after(() => {
    Object.defineProperty(process, 'platform', { value: originalPlatform });
    // Construction on Linux starts the shared pointer poll and z-order keep-alive loops.
    stopLinuxOverlayPointerInteractionLoop();
    stopLinuxOverlayZOrderKeepAliveLoop();
  });

  const mouseEvents: MouseEventsCall[] = [];
  const clearedLayers: string[] = [];
  const mainWindow = {
    isDestroyed: () => false,
    isVisible: () => true,
    isFocused: () => true,
    getBounds: () => ({ x: 0, y: 0, width: 1920, height: 1080 }),
    setShape: () => {},
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      mouseEvents.push({ ignore, options });
    },
  } as unknown as BrowserWindow;
  const noop = () => {};
  const deps: VisibleOverlayInteractionRuntimeDeps = {
    screen: {
      getCursorScreenPoint: () => ({ x: 5, y: 5 }),
      screenToDipPoint: (point) => point,
    },
    overlayManager: {
      getMainWindow: () => mainWindow,
      getVisibleOverlayVisible: () => true,
    },
    overlayContentMeasurementStore: {
      clear: (layer) => {
        clearedLayers.push(layer);
      },
      getLatestByLayer: () => null,
    },
    logger: { info: noop, warn: noop, debug: noop },
    updateVisibleOverlayVisibility: noop,
    getModalInputExclusive: () => false,
    getStatsOverlayVisible: () => false,
    setStatsOverlayVisible: noop,
    getWindowTracker: () => null,
    setWindowTracker: noop,
    setTrackerNotReadyWarningShown: noop,
    getMpvSocketPath: () => '/tmp/subminer-interaction-test.sock',
    getBackendOverride: () => null,
    getInitialArgs: () => null,
    getOverlayRuntimeInitialized: () => true,
    getLinuxVisibleOverlayWindowMode: () => 'managed',
    setLinuxVisibleOverlayOwnerBindingKey: noop,
    bindVisibleOverlayToTrackedX11Window: noop,
    updateVisibleOverlayBounds: noop,
    refreshCurrentSubtitle: noop,
    getOverlayWindows: () => [mainWindow],
    syncOverlayShortcuts: noop,
    resetLastOverlayWindowGeometry: noop,
    enforceOverlayLayerOrder: noop,
    getOverlayForegroundSeparateWindows: () => [],
  };
  const runtime = createVisibleOverlayInteractionRuntime(deps);
  return { runtime, mouseEvents, clearedLayers };
}

test('resetting visible overlay input clears the visible content measurement', (t) => {
  const { runtime, clearedLayers } = setup(t, 'darwin');

  runtime.resetVisibleOverlayInputState();

  assert.deepEqual(clearedLayers, ['visible']);
});

// With no subtitle under the cursor the poll wants passthrough. An input reset leaves the
// window's real state unknown, so the next poll must reapply it even though the desired state
// matches the remembered one. Hiding resets all input; showing resets only the startup primer.
for (const reset of [
  'resetVisibleOverlayInputState',
  'resetLinuxVisibleOverlayStartupInputPrimer',
] as const) {
  test(`the Linux pointer poll reapplies passive passthrough after ${reset}`, (t) => {
    const { runtime, mouseEvents } = setup(t, 'linux');
    const passthrough: MouseEventsCall = { ignore: true, options: { forward: true } };

    runtime.tickLinuxOverlayPointerInteractionNow();
    runtime.tickLinuxOverlayPointerInteractionNow();
    assert.deepEqual(mouseEvents, [passthrough]);

    runtime[reset]();
    runtime.tickLinuxOverlayPointerInteractionNow();
    runtime.tickLinuxOverlayPointerInteractionNow();
    assert.deepEqual(mouseEvents, [passthrough, passthrough]);
  });
}

test('the Linux startup input grace makes the overlay interactive on the next poll', (t) => {
  const { runtime, mouseEvents } = setup(t, 'linux');
  runtime.tickLinuxOverlayPointerInteractionNow();

  runtime.startLinuxVisibleOverlayStartupInputGrace();
  runtime.tickLinuxOverlayPointerInteractionNow();

  assert.deepEqual(mouseEvents.at(-1), { ignore: false, options: undefined });
  assert.equal(runtime.getVisibleOverlayInteractionActive(), true);
});

test('the Linux startup input grace reasserts input on an overlay that is already interactive', (t) => {
  const { runtime, mouseEvents } = setup(t, 'linux');
  const interactive: MouseEventsCall = { ignore: false, options: undefined };
  runtime.setLinuxOverlayInteractiveHint(true);
  runtime.tickLinuxOverlayPointerInteractionNow();
  runtime.tickLinuxOverlayPointerInteractionNow();
  assert.deepEqual(mouseEvents, [interactive]);

  runtime.startLinuxVisibleOverlayStartupInputGrace();
  runtime.tickLinuxOverlayPointerInteractionNow();

  assert.deepEqual(mouseEvents, [interactive, interactive]);
});
