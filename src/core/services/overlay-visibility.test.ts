import assert from 'node:assert/strict';
import test from 'node:test';

import { setVisibleOverlayVisible, updateVisibleOverlayVisibility } from './overlay-visibility';

type WindowTrackerStub = {
  isTracking: () => boolean;
  getGeometry: () => { x: number; y: number; width: number; height: number } | null;
};

function createMainWindowRecorder() {
  const calls: string[] = [];
  const window = {
    isDestroyed: () => false,
    hide: () => {
      calls.push('hide');
    },
    show: () => {
      calls.push('show');
    },
    focus: () => {
      calls.push('focus');
    },
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      calls.push(`mouse-ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`);
    },
  };

  return { window, calls };
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

test('untracked non-macOS overlay keeps fallback visible behavior when no tracker exists', () => {
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
  assert.ok(calls.includes('show'));
  assert.ok(calls.includes('focus'));
  assert.ok(!calls.includes('osd'));
});

test('Windows visible overlay stays click-through and does not steal focus while tracked', () => {
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
    isWindowsPlatform: true,
  } as never);

  assert.ok(calls.includes('mouse-ignore:true:forward'));
  assert.ok(calls.includes('show'));
  assert.ok(!calls.includes('focus'));
});

test('macOS tracked visible overlay stays visible without passively stealing focus', () => {
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

  assert.ok(calls.includes('mouse-ignore:false:plain'));
  assert.ok(calls.includes('show'));
  assert.ok(!calls.includes('focus'));
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
  assert.ok(calls.includes('show'));
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

test('macOS loading OSD can show again after overlay is hidden and retried', () => {
  const { window, calls } = createMainWindowRecorder();
  const osdMessages: string[] = [];
  let trackerWarning = false;

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
  } as never);

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
  } as never);

  assert.deepEqual(osdMessages, ['Overlay loading...', 'Overlay loading...']);
});
