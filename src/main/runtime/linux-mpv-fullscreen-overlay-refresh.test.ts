import assert from 'node:assert/strict';
import test from 'node:test';
import {
  clearLinuxMpvFullscreenOverlayRefreshTimeouts,
  updateLinuxMpvFullscreenOverlayRefreshBurst,
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
    scheduleLinuxVisibleOverlayFullscreenRefreshBurst(true, {
      overlayManager: {
        getMainWindow: () =>
          ({
            hide: () => calls.push('hide'),
            isFullScreen: () => false,
            isDestroyed: () => false,
            isVisible: () => true,
            setFullScreen: (fullscreen: boolean) => calls.push(`fullscreen:${fullscreen}`),
            setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) =>
              calls.push(
                `mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`,
              ),
            showInactive: () => calls.push('showInactive'),
          }) as never,
        getVisibleOverlayVisible: () => true,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => calls.push('updateVisibleOverlayVisibility'),
      },
      syncVisibleOverlayMpvFullscreenMode: (fullscreen: boolean) =>
        calls.push(`sync-overlay-mode:${fullscreen}`),
      ensureOverlayWindowLevel: () => calls.push('ensureOverlayWindowLevel'),
    });

    const deadline = Date.now() + 200;
    while (!calls.includes('updateVisibleOverlayVisibility') && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 5));
    }

    assert.ok(calls.includes('updateVisibleOverlayVisibility'));
    assert.ok(calls.includes('sync-overlay-mode:true'));
    assert.ok(!calls.includes('fullscreen:true'));
    assert.ok(calls.includes('hide'));
    assert.ok(calls.includes('showInactive'));
    assert.ok(calls.includes('mouse-ignore:true:forward'));
    assert.ok(calls.includes('ensureOverlayWindowLevel'));
  } finally {
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('linux mpv fullscreen overlay refresh remembers mode even when overlay is hidden', async () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  const calls: string[] = [];

  try {
    scheduleLinuxVisibleOverlayFullscreenRefreshBurst(true, {
      overlayManager: {
        getMainWindow: () => null,
        getVisibleOverlayVisible: () => false,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => calls.push('updateVisibleOverlayVisibility'),
      },
      syncVisibleOverlayMpvFullscreenMode: (fullscreen: boolean) =>
        calls.push(`sync-overlay-mode:${fullscreen}`),
      ensureOverlayWindowLevel: () => calls.push('ensureOverlayWindowLevel'),
    });

    const deadline = Date.now() + 200;
    while (!calls.includes('sync-overlay-mode:true') && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 5));
    }

    assert.ok(calls.includes('sync-overlay-mode:true'));
    assert.ok(!calls.includes('updateVisibleOverlayVisibility'));
    assert.ok(!calls.includes('ensureOverlayWindowLevel'));
  } finally {
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('linux mpv fullscreen overlay refresh updates mode without hide/show when fullscreen exits', async () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  const calls: string[] = [];

  try {
    const deps = {
      overlayManager: {
        getMainWindow: () =>
          ({
            hide: () => calls.push('hide'),
            isFullScreen: () => true,
            isDestroyed: () => false,
            isVisible: () => true,
            setFullScreen: (fullscreen: boolean) => calls.push(`fullscreen:${fullscreen}`),
            setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) =>
              calls.push(
                `mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`,
              ),
            showInactive: () => calls.push('showInactive'),
          }) as never,
        getVisibleOverlayVisible: () => true,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => calls.push('updateVisibleOverlayVisibility'),
      },
      syncVisibleOverlayMpvFullscreenMode: (fullscreen: boolean) =>
        calls.push(`sync-overlay-mode:${fullscreen}`),
      ensureOverlayWindowLevel: () => calls.push('ensureOverlayWindowLevel'),
    };

    const cancel = updateLinuxMpvFullscreenOverlayRefreshBurst(true, deps, null);
    const nextCancel = updateLinuxMpvFullscreenOverlayRefreshBurst(false, deps, cancel);

    await new Promise((resolve) => setTimeout(resolve, 80));

    assert.equal(typeof nextCancel, 'function');
    assert.ok(calls.includes('updateVisibleOverlayVisibility'));
    assert.ok(calls.includes('sync-overlay-mode:false'));
    assert.ok(!calls.includes('fullscreen:false'));
    assert.equal(calls.includes('hide'), false);
    assert.equal(calls.includes('showInactive'), false);
    assert.equal(calls.includes('mouse-ignore:true:forward'), false);
    assert.equal(calls.includes('ensureOverlayWindowLevel'), false);
  } finally {
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('linux mpv fullscreen overlay refresh restores click-through after restacking', async () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  const calls: string[] = [];

  try {
    scheduleLinuxVisibleOverlayFullscreenRefreshBurst(true, {
      overlayManager: {
        getMainWindow: () =>
          ({
            hide: () => calls.push('hide'),
            isFullScreen: () => false,
            isDestroyed: () => false,
            isVisible: () => true,
            setFullScreen: (fullscreen: boolean) => calls.push(`fullscreen:${fullscreen}`),
            setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) =>
              calls.push(
                `mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`,
              ),
            showInactive: () => calls.push('showInactive'),
          }) as never,
        getVisibleOverlayVisible: () => true,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => calls.push('updateVisibleOverlayVisibility'),
      },
      syncVisibleOverlayMpvFullscreenMode: (fullscreen: boolean) =>
        calls.push(`sync-overlay-mode:${fullscreen}`),
      ensureOverlayWindowLevel: () => calls.push('ensureOverlayWindowLevel'),
    });

    const deadline = Date.now() + 200;
    while (!calls.includes('mouse-ignore:true:forward') && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 5));
    }

    const showIndex = calls.indexOf('showInactive');
    const passthroughIndex = calls.indexOf('mouse-ignore:true:forward');
    const levelIndex = calls.indexOf('ensureOverlayWindowLevel');
    const syncIndex = calls.indexOf('sync-overlay-mode:true');

    assert.ok(syncIndex >= 0);
    assert.ok(showIndex >= 0);
    assert.ok(syncIndex < showIndex);
    assert.ok(passthroughIndex > showIndex);
    assert.ok(levelIndex > passthroughIndex);
  } finally {
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('linux mpv fullscreen overlay refresh preserves active subtitle interaction after restacking', async () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  const calls: string[] = [];

  try {
    scheduleLinuxVisibleOverlayFullscreenRefreshBurst(true, {
      overlayManager: {
        getMainWindow: () =>
          ({
            hide: () => calls.push('hide'),
            isFullScreen: () => false,
            isDestroyed: () => false,
            isVisible: () => true,
            setFullScreen: (fullscreen: boolean) => calls.push(`fullscreen:${fullscreen}`),
            setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) =>
              calls.push(
                `mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`,
              ),
            showInactive: () => calls.push('showInactive'),
          }) as never,
        getVisibleOverlayVisible: () => true,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => calls.push('updateVisibleOverlayVisibility'),
      },
      syncVisibleOverlayMpvFullscreenMode: (fullscreen: boolean) =>
        calls.push(`sync-overlay-mode:${fullscreen}`),
      getOverlayInteractionActive: () => true,
      ensureOverlayWindowLevel: () => calls.push('ensureOverlayWindowLevel'),
    });

    const deadline = Date.now() + 200;
    while (!calls.includes('mouse-ignore:false:plain') && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 5));
    }

    const showIndex = calls.indexOf('showInactive');
    const interactiveIndex = calls.indexOf('mouse-ignore:false:plain');

    assert.ok(showIndex >= 0);
    assert.ok(interactiveIndex > showIndex);
    assert.equal(calls.includes('mouse-ignore:true:forward'), false);
  } finally {
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});
