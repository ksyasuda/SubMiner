import assert from 'node:assert/strict';
import test from 'node:test';
import {
  resolveLinuxVisibleOverlayWindowModeAction,
  shouldExitFullscreenOverrideForTrackedGeometry,
} from './linux-visible-overlay-window-mode';

test('linux overlay mode sync records fullscreen without creating a hidden overlay', () => {
  assert.deepEqual(
    resolveLinuxVisibleOverlayWindowModeAction({
      currentMode: 'managed',
      fullscreen: true,
      hasLiveWindow: false,
      visibleOverlayVisible: false,
    }),
    {
      nextMode: 'fullscreen-override',
      shouldCreateWindow: false,
      shouldDestroyCurrentWindow: false,
      shouldRefreshVisibleOverlay: false,
      createWindowTiming: 'none',
    },
  );
});

test('linux overlay mode sync destroys stale hidden window without replacing it', () => {
  assert.deepEqual(
    resolveLinuxVisibleOverlayWindowModeAction({
      currentMode: 'managed',
      fullscreen: true,
      hasLiveWindow: true,
      visibleOverlayVisible: false,
    }),
    {
      nextMode: 'fullscreen-override',
      shouldCreateWindow: false,
      shouldDestroyCurrentWindow: true,
      shouldRefreshVisibleOverlay: false,
      createWindowTiming: 'none',
    },
  );
});

test('linux overlay mode sync replaces visible window when fullscreen mode changes', () => {
  assert.deepEqual(
    resolveLinuxVisibleOverlayWindowModeAction({
      currentMode: 'managed',
      fullscreen: true,
      hasLiveWindow: true,
      visibleOverlayVisible: true,
    }),
    {
      nextMode: 'fullscreen-override',
      shouldCreateWindow: true,
      shouldDestroyCurrentWindow: true,
      shouldRefreshVisibleOverlay: true,
      createWindowTiming: 'after-current-destroyed',
    },
  );
});

test('linux overlay mode sync creates correct visible window when none exists', () => {
  assert.deepEqual(
    resolveLinuxVisibleOverlayWindowModeAction({
      currentMode: 'fullscreen-override',
      fullscreen: true,
      hasLiveWindow: false,
      visibleOverlayVisible: true,
    }),
    {
      nextMode: 'fullscreen-override',
      shouldCreateWindow: true,
      shouldDestroyCurrentWindow: false,
      shouldRefreshVisibleOverlay: true,
      createWindowTiming: 'now',
    },
  );
});

test('linux overlay mode sync no-ops when live window already matches mode', () => {
  assert.deepEqual(
    resolveLinuxVisibleOverlayWindowModeAction({
      currentMode: 'fullscreen-override',
      fullscreen: true,
      hasLiveWindow: true,
      visibleOverlayVisible: true,
    }),
    {
      nextMode: 'fullscreen-override',
      shouldCreateWindow: false,
      shouldDestroyCurrentWindow: false,
      shouldRefreshVisibleOverlay: false,
      createWindowTiming: 'none',
    },
  );
});

test('linux overlay mode exits fullscreen override when tracked geometry is windowed', () => {
  assert.equal(
    shouldExitFullscreenOverrideForTrackedGeometry({
      currentMode: 'fullscreen-override',
      trackedFullscreen: true,
      geometry: { x: 420, y: 90, width: 1280, height: 720 },
      displayBounds: { x: 0, y: 0, width: 2560, height: 1440 },
    }),
    true,
  );
});
