import assert from 'node:assert/strict';
import test from 'node:test';

import { OVERLAY_WINDOW_CONTENT_READY_FLAG } from './overlay-window-flags';
import { setVisibleOverlayVisible, updateVisibleOverlayVisibility } from './overlay-visibility';

type WindowTrackerStub = {
  isTracking: () => boolean;
  getGeometry: () => { x: number; y: number; width: number; height: number } | null;
  isTargetWindowFocused?: () => boolean;
  isTargetWindowMinimized?: () => boolean;
};

function createMainWindowRecorder(options: { emitShowImmediately?: boolean } = {}) {
  const emitShowImmediately = options.emitShowImmediately ?? true;
  const calls: string[] = [];
  const listeners = new Map<string, Array<() => void>>();
  let visible = false;
  let focused = false;
  let opacity = 1;
  let contentReady = true;
  const emit = (event: string): void => {
    const handlers = listeners.get(event) ?? [];
    listeners.delete(event);
    for (const handler of handlers) {
      handler();
    }
  };
  const emitShow = (): void => {
    visible = true;
    emit('show');
  };
  const window = {
    webContents: {},
    isDestroyed: () => false,
    isVisible: () => visible,
    isFocused: () => focused,
    once: (event: string, handler: () => void) => {
      listeners.set(event, [...(listeners.get(event) ?? []), handler]);
    },
    hide: () => {
      visible = false;
      focused = false;
      calls.push('hide');
    },
    show: () => {
      calls.push('show');
      if (emitShowImmediately) {
        emitShow();
      }
    },
    showInactive: () => {
      calls.push('show-inactive');
      if (emitShowImmediately) {
        emitShow();
      }
    },
    focus: () => {
      focused = true;
      calls.push('focus');
    },
    setAlwaysOnTop: (flag: boolean) => {
      calls.push(`always-on-top:${flag}`);
    },
    setVisibleOnAllWorkspaces: (flag: boolean, options?: { visibleOnFullScreen?: boolean }) => {
      calls.push(
        `all-workspaces:${flag}:${options?.visibleOnFullScreen === true ? 'fullscreen' : 'plain'}`,
      );
    },
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      calls.push(`mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`);
    },
    setOpacity: (nextOpacity: number) => {
      opacity = nextOpacity;
      calls.push(`opacity:${nextOpacity}`);
    },
    moveTop: () => {
      calls.push('move-top');
    },
  };
  (
    window as {
      [OVERLAY_WINDOW_CONTENT_READY_FLAG]?: boolean;
    }
  )[OVERLAY_WINDOW_CONTENT_READY_FLAG] = contentReady;

  return {
    window,
    calls,
    getOpacity: () => opacity,
    emitShow,
    setContentReady: (nextContentReady: boolean) => {
      contentReady = nextContentReady;
      (
        window as {
          [OVERLAY_WINDOW_CONTENT_READY_FLAG]?: boolean;
        }
      )[OVERLAY_WINDOW_CONTENT_READY_FLAG] = contentReady;
    },
    setFocused: (nextFocused: boolean) => {
      focused = nextFocused;
    },
  };
}

test('macOS keeps visible overlay hidden while tracker is not ready and emits one loading OSD', () => {
  const { window, calls } = createMainWindowRecorder();
  let trackerWarning = false;
  const osdMessages: string[] = [];
  const tracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
  };

  const run = () =>
    updateVisibleOverlayVisibility({
      visibleOverlayVisible: true,
      mainWindow: window as never,
      windowTracker: tracker as never,
      trackerNotReadyWarningShown: trackerWarning,
      setTrackerNotReadyWarningShown: (shown: boolean) => {
        trackerWarning = shown;
      },
      updateVisibleOverlayBounds: () => {
        calls.push('update-bounds');
      },
      ensureOverlayWindowLevel: () => {
        calls.push('ensure-level');
      },
      syncPrimaryOverlayWindowLayer: () => {
        calls.push('sync-layer');
      },
      enforceOverlayLayerOrder: () => {
        calls.push('enforce-order');
      },
      syncOverlayShortcuts: () => {
        calls.push('sync-shortcuts');
      },
      isMacOSPlatform: true,
      showOverlayLoadingOsd: (message: string) => {
        osdMessages.push(message);
      },
    } as never);

  run();
  run();

  assert.equal(trackerWarning, true);
  assert.deepEqual(osdMessages, ['Overlay loading...']);
  assert.ok(calls.includes('hide'));
  assert.ok(!calls.includes('show'));
});

test('tracked non-macOS overlay stays hidden while tracker is not ready', () => {
  const { window, calls } = createMainWindowRecorder();
  let trackerWarning = false;
  const tracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    showOverlayLoadingOsd: () => {
      calls.push('osd');
    },
    resolveFallbackBounds: () => ({ x: 12, y: 24, width: 640, height: 360 }),
  } as never);

  assert.equal(trackerWarning, true);
  assert.ok(calls.includes('hide'));
  assert.ok(!calls.includes('update-bounds'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('focus'));
  assert.ok(!calls.includes('osd'));
});

test('non-native passive overlay stays click-through after subsequent visibility updates', () => {
  const { window, calls } = createMainWindowRecorder();

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: false,
    overlayInteractionActive: false,
    showOverlayLoadingOsd: () => {},
    resolveFallbackBounds: () => ({ x: 12, y: 24, width: 640, height: 360 }),
  } as never);
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: false,
    overlayInteractionActive: false,
    showOverlayLoadingOsd: () => {},
    resolveFallbackBounds: () => ({ x: 12, y: 24, width: 640, height: 360 }),
  } as never);

  assert.equal(calls.includes('mouse-ignore:false:plain'), false);
  assert.ok(calls.includes('mouse-ignore:true:forward'));
});

test('suspended visible overlay hides without refreshing bounds or z-order', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  window.show();
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    suspendVisibleOverlay: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('always-on-top:false'));
  assert.ok(calls.includes('hide'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('update-bounds'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('focus'));
});

test('untracked non-macOS overlay shows passively when no tracker exists', () => {
  const { window, calls } = createMainWindowRecorder();
  let trackerWarning = false;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: null,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    showOverlayLoadingOsd: () => {
      calls.push('osd');
    },
    resolveFallbackBounds: () => ({ x: 12, y: 24, width: 640, height: 360 }),
  } as never);

  assert.equal(trackerWarning, false);
  assert.ok(calls.includes('show-inactive'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('focus'));
  assert.ok(!calls.includes('osd'));
});

test('passive Linux visible overlay does not take keyboard focus', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('show-inactive'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('focus'));
});

test('tracked non-macOS overlay reapplies bounds after first show', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: false,
  } as never);

  assert.deepEqual(
    calls.filter((call) => call === 'update-bounds' || call === 'show-inactive'),
    ['update-bounds', 'show-inactive', 'update-bounds'],
  );
});

test('tracked non-macOS overlay queues only one first-show bounds refresh', () => {
  const { window, calls, emitShow } = createMainWindowRecorder({ emitShowImmediately: false });
  let width = 1280;
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width, height: 720 }),
  };
  const run = () =>
    updateVisibleOverlayVisibility({
      visibleOverlayVisible: true,
      mainWindow: window as never,
      windowTracker: tracker as never,
      trackerNotReadyWarningShown: false,
      setTrackerNotReadyWarningShown: () => {},
      updateVisibleOverlayBounds: (geometry: { width: number }) => {
        calls.push(`update-bounds:${geometry.width}`);
      },
      ensureOverlayWindowLevel: () => {
        calls.push('ensure-level');
      },
      syncPrimaryOverlayWindowLayer: () => {
        calls.push('sync-layer');
      },
      enforceOverlayLayerOrder: () => {
        calls.push('enforce-order');
      },
      syncOverlayShortcuts: () => {
        calls.push('sync-shortcuts');
      },
      isMacOSPlatform: false,
      isWindowsPlatform: false,
    } as never);

  run();
  width = 1440;
  run();
  emitShow();

  assert.deepEqual(
    calls.filter((call) => call.startsWith('update-bounds:')),
    ['update-bounds:1280', 'update-bounds:1440', 'update-bounds:1440'],
  );
});

test('Windows visible overlay stays click-through and binds to mpv while tracked', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(calls.includes('opacity:0'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('show-inactive'));
  assert.ok(calls.includes('sync-windows-z-order'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
  assert.ok(!calls.includes('focus'));
});

test('Windows visible overlay restores opacity after the deferred reveal delay', async () => {
  const { window, calls, getOpacity } = createMainWindowRecorder();
  let syncWindowsZOrderCalls = 0;
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      syncWindowsZOrderCalls += 1;
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.equal(getOpacity(), 0);
  assert.equal(syncWindowsZOrderCalls, 1);
  await new Promise<void>((resolve) => setTimeout(resolve, 60));
  assert.equal(getOpacity(), 1);
  assert.equal(syncWindowsZOrderCalls, 2);
  assert.ok(calls.includes('opacity:1'));
});

test('Windows visible overlay waits for content-ready before first reveal', () => {
  const { window, calls, setContentReady } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };
  setContentReady(false);

  const run = () =>
    updateVisibleOverlayVisibility({
      visibleOverlayVisible: true,
      mainWindow: window as never,
      windowTracker: tracker as never,
      trackerNotReadyWarningShown: false,
      setTrackerNotReadyWarningShown: () => {},
      updateVisibleOverlayBounds: () => {
        calls.push('update-bounds');
      },
      ensureOverlayWindowLevel: () => {
        calls.push('ensure-level');
      },
      syncWindowsOverlayToMpvZOrder: () => {
        calls.push('sync-windows-z-order');
      },
      syncPrimaryOverlayWindowLayer: () => {
        calls.push('sync-layer');
      },
      enforceOverlayLayerOrder: () => {
        calls.push('enforce-order');
      },
      syncOverlayShortcuts: () => {
        calls.push('sync-shortcuts');
      },
      isMacOSPlatform: false,
      isWindowsPlatform: true,
    } as never);

  run();

  assert.ok(!calls.includes('show-inactive'));
  assert.ok(!calls.includes('show'));

  setContentReady(true);
  run();

  assert.ok(calls.includes('show-inactive'));
});

test('tracked Windows overlay refresh rebinds while already visible', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('sync-windows-z-order'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(calls.includes('sync-shortcuts'));
});

test('forced passthrough still reapplies while visible on Windows', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
    forceMousePassthrough: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(calls.includes('sync-windows-z-order'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
});

test('forced passthrough still shows tracked overlay while bound to mpv on Windows', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
    forceMousePassthrough: true,
  } as never);

  assert.ok(calls.includes('show-inactive'));
  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(calls.includes('sync-windows-z-order'));
});

test('forced mouse passthrough keeps macOS tracked overlay above active mpv', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => true,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
    forceMousePassthrough: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(!calls.includes('always-on-top:false'));
});

test('forced mouse passthrough still hides macOS tracked overlay when mpv loses foreground', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  window.show();
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
    forceMousePassthrough: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('always-on-top:false'));
  assert.ok(calls.includes('all-workspaces:false:plain'));
  assert.ok(calls.includes('hide'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
});

test('tracked Windows overlay rebinds without hiding when tracker focus changes', () => {
  const { window, calls } = createMainWindowRecorder();
  let focused = true;
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => focused,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  calls.length = 0;
  focused = false;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('sync-windows-z-order'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
  assert.ok(!calls.includes('show'));
});

test('tracked Windows overlay stays interactive while the overlay window itself is focused', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  calls.length = 0;
  setFocused(true);

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:false:plain'));
  assert.ok(calls.includes('sync-windows-z-order'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
});

test('tracked Windows overlay reshows click-through even if focus state is stale after a modal closes', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  calls.length = 0;
  window.hide();
  calls.length = 0;
  setFocused(true);

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('show-inactive'));
  assert.ok(!calls.includes('show'));
});

test('tracked Windows overlay binds above mpv even when tracker focus lags', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('sync-windows-z-order'));
  assert.ok(!calls.includes('ensure-level'));
});

test('visible overlay stays hidden while a modal window is active', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    modalActive: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('hide'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('update-bounds'));
});

test('macOS tracked visible overlay starts click-through without passively stealing focus', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('show-inactive'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('focus'));
});

test('macOS tracked visible overlay remains click-through even if the overlay had focus', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => true,
  };

  setFocused(true);

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(!calls.includes('focus'));
});

test('macOS keeps active mpv overlay visible and click-through during tracker refresh', () => {
  const { window, calls } = createMainWindowRecorder();
  const osdMessages: string[] = [];
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => true,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {
      calls.push('tracker-warning');
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
    showOverlayLoadingOsd: (message: string) => {
      osdMessages.push(message);
    },
  } as never);

  assert.ok(calls.includes('update-bounds'));
  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('hide'));
  assert.deepEqual(osdMessages, []);
});

test('macOS tracked overlay hides when mpv loses foreground', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  window.show();
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('update-bounds'));
  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('always-on-top:false'));
  assert.ok(calls.includes('all-workspaces:false:plain'));
  assert.ok(calls.includes('hide'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
  assert.ok(!calls.includes('focus'));
  assert.ok(!calls.includes('show'));
});

test('macOS keeps visible overlay stable while probing frontmost app after overlay blur', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
    isTargetWindowMinimized: () => false,
  };

  window.show();
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
    macOSForegroundProbeActive: true,
  } as never);

  assert.ok(calls.includes('update-bounds'));
  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('hide'));
});

test('macOS keeps tracked overlay visible while overlay interaction is active after mpv loses foreground', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  window.show();
  setFocused(false);
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    overlayInteractionActive: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('update-bounds'));
  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:false:plain'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('hide'));
});

test('macOS lets an active overlay receive mouse input instead of forcing passthrough', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  window.show();
  setFocused(false);
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    overlayInteractionActive: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('mouse-ignore:false:plain'));
  assert.ok(!calls.includes('mouse-ignore:true:forward'));
  assert.ok(!calls.includes('hide'));
});

test('macOS focuses an active overlay so lookup trigger keys reach it', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
  };

  window.show();
  setFocused(false);
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    overlayInteractionActive: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
  } as never);

  assert.ok(calls.includes('mouse-ignore:false:plain'));
  assert.ok(calls.includes('focus'));
  assert.ok(!calls.includes('hide'));
});

test('macOS tracked overlay passively reappears when mpv regains foreground', () => {
  const { window, calls } = createMainWindowRecorder();
  let targetFocused = false;
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => targetFocused,
  };

  window.show();
  calls.length = 0;

  const run = () =>
    updateVisibleOverlayVisibility({
      visibleOverlayVisible: true,
      mainWindow: window as never,
      windowTracker: tracker as never,
      trackerNotReadyWarningShown: false,
      setTrackerNotReadyWarningShown: () => {},
      updateVisibleOverlayBounds: () => {
        calls.push('update-bounds');
      },
      ensureOverlayWindowLevel: () => {
        calls.push('ensure-level');
      },
      syncPrimaryOverlayWindowLayer: () => {
        calls.push('sync-layer');
      },
      enforceOverlayLayerOrder: () => {
        calls.push('enforce-order');
      },
      syncOverlayShortcuts: () => {
        calls.push('sync-shortcuts');
      },
      isMacOSPlatform: true,
      isWindowsPlatform: false,
    } as never);

  run();
  assert.ok(calls.includes('hide'));

  calls.length = 0;
  targetFocused = true;
  run();

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('show-inactive'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('focus'));
});

test('macOS preserves an already visible active mpv overlay while tracker is temporarily not ready', () => {
  const { window, calls } = createMainWindowRecorder();
  const osdMessages: string[] = [];
  let trackerWarning = false;
  const tracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
    isTargetWindowFocused: () => true,
  };

  window.show();
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
      calls.push(`tracker-warning:${shown}`);
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
    showOverlayLoadingOsd: (message: string) => {
      osdMessages.push(message);
    },
  } as never);

  assert.equal(trackerWarning, false);
  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('hide'));
  assert.deepEqual(osdMessages, []);
});

test('forced mouse passthrough keeps macOS tracked overlay passive while visible', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    isWindowsPlatform: false,
    forceMousePassthrough: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('show-inactive'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('focus'));
});

test('Windows keeps visible overlay hidden while tracker is not ready', () => {
  const { window, calls } = createMainWindowRecorder();
  let trackerWarning = false;
  const tracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
    resolveFallbackBounds: () => ({ x: 12, y: 24, width: 640, height: 360 }),
  } as never);

  assert.equal(trackerWarning, true);
  assert.ok(calls.includes('hide'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('update-bounds'));
});

test('Windows preserves visible overlay and rebinds to mpv while tracker transiently loses a non-minimized window', () => {
  const { window, calls } = createMainWindowRecorder();
  let tracking = true;
  const tracker: WindowTrackerStub = {
    isTracking: () => tracking,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => false,
    isTargetWindowMinimized: () => false,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  calls.length = 0;
  tracking = false;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(!calls.includes('hide'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('move-top'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('sync-windows-z-order'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(calls.includes('sync-shortcuts'));
});

test('Windows hides the visible overlay when the tracked window is minimized', () => {
  const { window, calls } = createMainWindowRecorder();
  let tracking = true;
  const tracker: WindowTrackerStub = {
    isTracking: () => tracking,
    getGeometry: () => (tracking ? { x: 0, y: 0, width: 1280, height: 720 } : null),
    isTargetWindowMinimized: () => !tracking,
  };

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  calls.length = 0;
  tracking = false;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncWindowsOverlayToMpvZOrder: () => {
      calls.push('sync-windows-z-order');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: false,
    isWindowsPlatform: true,
  } as never);

  assert.ok(calls.includes('hide'));
  assert.ok(!calls.includes('sync-windows-z-order'));
});

test('macOS keeps visible overlay hidden while tracker is not initialized yet', () => {
  const { window, calls } = createMainWindowRecorder();
  let trackerWarning = false;
  const osdMessages: string[] = [];

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: null,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    showOverlayLoadingOsd: (message: string) => {
      osdMessages.push(message);
    },
  } as never);

  assert.equal(trackerWarning, true);
  assert.deepEqual(osdMessages, ['Overlay loading...']);
  assert.ok(calls.includes('hide'));
  assert.ok(!calls.includes('show'));
  assert.ok(!calls.includes('update-bounds'));
});

test('macOS preserves visible overlay during transient tracker loss with retained geometry', () => {
  const { window, calls } = createMainWindowRecorder();
  const osdMessages: string[] = [];
  let trackerWarning = false;
  let tracking = true;
  const tracker: WindowTrackerStub = {
    isTracking: () => tracking,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
    isTargetWindowFocused: () => true,
  };

  const run = () =>
    updateVisibleOverlayVisibility({
      visibleOverlayVisible: true,
      mainWindow: window as never,
      windowTracker: tracker as never,
      trackerNotReadyWarningShown: trackerWarning,
      setTrackerNotReadyWarningShown: (shown: boolean) => {
        trackerWarning = shown;
      },
      updateVisibleOverlayBounds: () => {
        calls.push('update-bounds');
      },
      ensureOverlayWindowLevel: () => {
        calls.push('ensure-level');
      },
      syncPrimaryOverlayWindowLayer: () => {
        calls.push('sync-layer');
      },
      enforceOverlayLayerOrder: () => {
        calls.push('enforce-order');
      },
      syncOverlayShortcuts: () => {
        calls.push('sync-shortcuts');
      },
      isMacOSPlatform: true,
      showOverlayLoadingOsd: (message: string) => {
        osdMessages.push(message);
      },
    } as never);

  run();
  calls.length = 0;
  tracking = false;

  run();

  assert.equal(trackerWarning, false);
  assert.deepEqual(osdMessages, []);
  assert.ok(calls.includes('update-bounds'));
  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('hide'));
  assert.ok(!calls.includes('show'));
});

test('macOS hides visible overlay during tracker loss after mpv loses foreground', () => {
  const { window, calls } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
    isTargetWindowFocused: () => false,
    isTargetWindowMinimized: () => false,
  };

  window.show();
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    showOverlayLoadingOsd: () => {
      calls.push('loading-osd');
    },
  } as never);

  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('always-on-top:false'));
  assert.ok(calls.includes('all-workspaces:false:plain'));
  assert.ok(calls.includes('hide'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('ensure-level'));
  assert.ok(!calls.includes('enforce-order'));
  assert.ok(!calls.includes('loading-osd'));
});

test('macOS keeps a focused overlay visible during tracker loss', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
    isTargetWindowFocused: () => false,
    isTargetWindowMinimized: () => false,
  };

  window.show();
  setFocused(true);
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    showOverlayLoadingOsd: () => {
      calls.push('loading-osd');
    },
  } as never);

  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('hide'));
  assert.ok(!calls.includes('loading-osd'));
});

test('macOS keeps an interactive overlay visible during tracker loss even when Electron focus drops', () => {
  const { window, calls, setFocused } = createMainWindowRecorder();
  const tracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
    isTargetWindowFocused: () => false,
    isTargetWindowMinimized: () => false,
  };

  window.show();
  setFocused(false);
  calls.length = 0;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    overlayInteractionActive: true,
    mainWindow: window as never,
    windowTracker: tracker as never,
    trackerNotReadyWarningShown: false,
    setTrackerNotReadyWarningShown: () => {},
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    showOverlayLoadingOsd: () => {
      calls.push('loading-osd');
    },
  } as never);

  assert.ok(calls.includes('sync-layer'));
  assert.ok(calls.includes('mouse-ignore:false:plain'));
  assert.ok(calls.includes('ensure-level'));
  assert.ok(calls.includes('enforce-order'));
  assert.ok(calls.includes('sync-shortcuts'));
  assert.ok(!calls.includes('always-on-top:false'));
  assert.ok(!calls.includes('hide'));
  assert.ok(!calls.includes('loading-osd'));
});

test('macOS suppresses immediate repeat loading OSD after tracker recovery until cooldown expires', () => {
  const { window } = createMainWindowRecorder();
  const osdMessages: string[] = [];
  let trackerWarning = false;
  let lastLoadingOsdAtMs: number | null = null;
  let nowMs = 1_000;
  const hiddenTracker: WindowTrackerStub = {
    isTracking: () => false,
    getGeometry: () => null,
  };
  const trackedTracker: WindowTrackerStub = {
    isTracking: () => true,
    getGeometry: () => ({ x: 0, y: 0, width: 1280, height: 720 }),
  };

  const run = (windowTracker: WindowTrackerStub) =>
    updateVisibleOverlayVisibility({
      visibleOverlayVisible: true,
      mainWindow: window as never,
      windowTracker: windowTracker as never,
      trackerNotReadyWarningShown: trackerWarning,
      setTrackerNotReadyWarningShown: (shown: boolean) => {
        trackerWarning = shown;
      },
      updateVisibleOverlayBounds: () => {},
      ensureOverlayWindowLevel: () => {},
      syncPrimaryOverlayWindowLayer: () => {},
      enforceOverlayLayerOrder: () => {},
      syncOverlayShortcuts: () => {},
      isMacOSPlatform: true,
      showOverlayLoadingOsd: (message: string) => {
        osdMessages.push(message);
      },
      shouldShowOverlayLoadingOsd: () =>
        lastLoadingOsdAtMs === null || nowMs - lastLoadingOsdAtMs >= 5_000,
      markOverlayLoadingOsdShown: () => {
        lastLoadingOsdAtMs = nowMs;
      },
    } as never);

  run(hiddenTracker);
  run(trackedTracker);

  nowMs = 2_000;
  run(hiddenTracker);
  run(trackedTracker);

  nowMs = 6_500;
  run(hiddenTracker);

  assert.deepEqual(osdMessages, ['Overlay loading...', 'Overlay loading...']);
});

test('setVisibleOverlayVisible does not mutate mpv subtitle visibility directly', () => {
  const calls: string[] = [];
  setVisibleOverlayVisible({
    visible: true,
    setVisibleOverlayVisibleState: (visible) => {
      calls.push(`state:${visible}`);
    },
    updateVisibleOverlayVisibility: () => {
      calls.push('update');
    },
  });

  assert.deepEqual(calls, ['state:true', 'update']);
});

test('macOS explicit hide resets loading OSD suppression before retry', () => {
  const { window, calls } = createMainWindowRecorder();
  const osdMessages: string[] = [];
  let trackerWarning = false;
  let lastLoadingOsdAtMs: number | null = null;
  let nowMs = 1_000;

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: null,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
      calls.push(`warn:${shown ? 'yes' : 'no'}`);
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    showOverlayLoadingOsd: (message: string) => {
      osdMessages.push(message);
    },
    shouldShowOverlayLoadingOsd: () =>
      lastLoadingOsdAtMs === null || nowMs - lastLoadingOsdAtMs >= 5_000,
    markOverlayLoadingOsdShown: () => {
      lastLoadingOsdAtMs = nowMs;
    },
    resetOverlayLoadingOsdSuppression: () => {
      lastLoadingOsdAtMs = null;
    },
  } as never);

  nowMs = 1_500;
  updateVisibleOverlayVisibility({
    visibleOverlayVisible: false,
    mainWindow: window as never,
    windowTracker: null,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
      calls.push(`warn:${shown ? 'yes' : 'no'}`);
    },
    updateVisibleOverlayBounds: () => {},
    ensureOverlayWindowLevel: () => {},
    syncPrimaryOverlayWindowLayer: () => {},
    enforceOverlayLayerOrder: () => {},
    syncOverlayShortcuts: () => {},
    isMacOSPlatform: true,
    showOverlayLoadingOsd: () => {},
    resetOverlayLoadingOsdSuppression: () => {
      lastLoadingOsdAtMs = null;
    },
  } as never);

  updateVisibleOverlayVisibility({
    visibleOverlayVisible: true,
    mainWindow: window as never,
    windowTracker: null,
    trackerNotReadyWarningShown: trackerWarning,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      trackerWarning = shown;
      calls.push(`warn:${shown ? 'yes' : 'no'}`);
    },
    updateVisibleOverlayBounds: () => {
      calls.push('update-bounds');
    },
    ensureOverlayWindowLevel: () => {
      calls.push('ensure-level');
    },
    syncPrimaryOverlayWindowLayer: () => {
      calls.push('sync-layer');
    },
    enforceOverlayLayerOrder: () => {
      calls.push('enforce-order');
    },
    syncOverlayShortcuts: () => {
      calls.push('sync-shortcuts');
    },
    isMacOSPlatform: true,
    showOverlayLoadingOsd: (message: string) => {
      osdMessages.push(message);
    },
    shouldShowOverlayLoadingOsd: () =>
      lastLoadingOsdAtMs === null || nowMs - lastLoadingOsdAtMs >= 5_000,
    markOverlayLoadingOsdShown: () => {
      lastLoadingOsdAtMs = nowMs;
    },
    resetOverlayLoadingOsdSuppression: () => {
      lastLoadingOsdAtMs = null;
    },
  } as never);

  assert.deepEqual(osdMessages, ['Overlay loading...', 'Overlay loading...']);
});
