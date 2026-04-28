import assert from 'node:assert/strict';
import test from 'node:test';
import {
  clearLinuxMpvFullscreenOverlayRefreshTimeouts,
  scheduleLinuxVisibleOverlayFullscreenRefreshBurst,
} from './linux-mpv-fullscreen-overlay-refresh';

test('linux mpv fullscreen overlay refresh burst schedules overlay refresh work on linux', async () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  const calls: string[] = [];

  try {
    scheduleLinuxVisibleOverlayFullscreenRefreshBurst({
      overlayManager: {
        getMainWindow: () =>
          ({
            hide: () => calls.push('hide'),
            isDestroyed: () => false,
            isVisible: () => true,
            showInactive: () => calls.push('showInactive'),
          }) as never,
        getVisibleOverlayVisible: () => true,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => calls.push('updateVisibleOverlayVisibility'),
      },
      ensureOverlayWindowLevel: () => calls.push('ensureOverlayWindowLevel'),
    });

    const deadline = Date.now() + 200;
    while (!calls.includes('updateVisibleOverlayVisibility') && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 5));
    }

    assert.ok(calls.includes('updateVisibleOverlayVisibility'));
    assert.ok(calls.includes('hide'));
    assert.ok(calls.includes('showInactive'));
    assert.ok(calls.includes('ensureOverlayWindowLevel'));
  } finally {
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});
