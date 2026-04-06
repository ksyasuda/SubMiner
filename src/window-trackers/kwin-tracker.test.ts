import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import * as vm from 'node:vm';
import { detectCompositor } from './index';
import {
  buildKWinBridgeScript,
  buildKWinTrackerPluginName,
  buildKWinTrackerServiceName,
  KWinWindowTracker,
  selectKWinMpvWindow,
  type KWinWindow,
} from './kwin-tracker';

const NativeWeakSet = globalThis.WeakSet;

type ScriptCallback = (...args: unknown[]) => void;

function createScriptSignal() {
  const callbacks: ScriptCallback[] = [];
  return {
    callbacks,
    connect(callback: ScriptCallback) {
      callbacks.push(callback);
    },
    emit(...args: unknown[]) {
      for (const callback of callbacks) {
        callback(...args);
      }
    },
  };
}

type ScriptSignal = ReturnType<typeof createScriptSignal>;

interface ScriptWindow {
  __weakSetUnsafe?: boolean;
  testId?: string;
  active: boolean;
  caption: string;
  managed: boolean;
  deleted: boolean;
  minimized: boolean;
  normalWindow: boolean;
  specialWindow: boolean;
  transient: boolean;
  popupWindow: boolean;
  outline: boolean;
  modal: boolean;
  pid: number;
  resourceClass: string;
  resourceName: string;
  keepAbove: boolean;
  visible: boolean;
  clientGeometry: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  frameGeometry: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  closed: ScriptSignal;
  clientGeometryChanged: ScriptSignal;
  frameGeometryChanged: ScriptSignal;
  outputChanged: ScriptSignal;
  windowClassChanged: ScriptSignal;
  windowShown: ScriptSignal;
  windowHidden: ScriptSignal;
  activeChanged: ScriptSignal;
}

interface ScriptWorkspace {
  windowList: () => ScriptWindow[];
  raiseWindow: (window: ScriptWindow) => void;
  windowAdded: ScriptSignal;
  windowRemoved: ScriptSignal;
  windowActivated: ScriptSignal;
  screensChanged: ScriptSignal;
  activeWindow?: ScriptWindow | null;
}

class GuardedWeakSet<T extends object> {
  private readonly inner = new NativeWeakSet<T>();

  has(value: T): boolean {
    this.assertSafe(value);
    return this.inner.has(value);
  }

  add(value: T): this {
    this.assertSafe(value);
    this.inner.add(value);
    return this;
  }

  private assertSafe(value: T): void {
    const candidate = value as { __weakSetUnsafe?: boolean; testId?: string };
    if (candidate.__weakSetUnsafe === true) {
      throw new Error(`unexpected WeakSet access for ${candidate.testId ?? 'window'}`);
    }
  }
}

function makeWindow(overrides: Partial<KWinWindow> = {}): KWinWindow {
  return {
    active: false,
    caption: 'mpv',
    minimized: false,
    normalWindow: true,
    pid: 100,
    resourceClass: 'mpv',
    resourceName: 'mpv',
    x: 10,
    y: 20,
    width: 1280,
    height: 720,
    ...overrides,
  };
}

function makeScriptWindow(overrides: Partial<ScriptWindow> = {}): ScriptWindow {
  return {
    active: false,
    caption: 'mpv',
    managed: true,
    deleted: false,
    minimized: false,
    normalWindow: true,
    specialWindow: false,
    transient: false,
    popupWindow: false,
    outline: false,
    modal: false,
    pid: 100,
    resourceClass: 'mpv',
    resourceName: 'mpv',
    keepAbove: false,
    visible: true,
    clientGeometry: {
      x: 10,
      y: 20,
      width: 1280,
      height: 720,
    },
    frameGeometry: {
      x: 10,
      y: 20,
      width: 1280,
      height: 720,
    },
    closed: createScriptSignal(),
    clientGeometryChanged: createScriptSignal(),
    frameGeometryChanged: createScriptSignal(),
    outputChanged: createScriptSignal(),
    windowClassChanged: createScriptSignal(),
    windowShown: createScriptSignal(),
    windowHidden: createScriptSignal(),
    activeChanged: createScriptSignal(),
    ...overrides,
  };
}

function makeOverlayScriptWindow(overrides: Partial<ScriptWindow> = {}): ScriptWindow {
  return makeScriptWindow({
    caption: 'SubMiner',
    pid: process.pid,
    resourceClass: 'subminer',
    resourceName: 'subminer',
    ...overrides,
  });
}

function trackWindowMutations(window: ScriptWindow): ScriptWindow & {
  frameGeometryAssignments: Array<{ x: number; y: number; width: number; height: number }>;
  minimizedAssignments: boolean[];
  keepAboveAssignments: boolean[];
} {
  let frameGeometryValue = { ...window.frameGeometry };
  let minimizedValue = window.minimized;
  let keepAboveValue = window.keepAbove;
  const frameGeometryAssignments: Array<{ x: number; y: number; width: number; height: number }> =
    [];
  const minimizedAssignments: boolean[] = [];
  const keepAboveAssignments: boolean[] = [];

  Object.defineProperty(window, 'frameGeometry', {
    configurable: true,
    enumerable: true,
    get: () => frameGeometryValue,
    set: (value) => {
      frameGeometryValue = {
        x: Number(value?.x || 0),
        y: Number(value?.y || 0),
        width: Number(value?.width || 0),
        height: Number(value?.height || 0),
      };
      frameGeometryAssignments.push(frameGeometryValue);
    },
  });

  Object.defineProperty(window, 'minimized', {
    configurable: true,
    enumerable: true,
    get: () => minimizedValue,
    set: (value) => {
      minimizedValue = value === true;
      minimizedAssignments.push(minimizedValue);
    },
  });

  Object.defineProperty(window, 'keepAbove', {
    configurable: true,
    enumerable: true,
    get: () => keepAboveValue,
    set: (value) => {
      keepAboveValue = value === true;
      keepAboveAssignments.push(keepAboveValue);
    },
  });

  return Object.assign(window, {
    frameGeometryAssignments,
    minimizedAssignments,
    keepAboveAssignments,
  });
}

function parseLastBridgePayload(payloads: string[]): { windows?: KWinWindow[] } {
  const payload = payloads.at(-1);
  assert.notEqual(payload, undefined);
  return JSON.parse(payload! as string) as { windows?: KWinWindow[] };
}

function runKWinBridgeScript(windows: ScriptWindow[]): {
  dbusPayloads: string[];
  workspace: ScriptWorkspace;
  activeWindowHistory: string[];
  raiseCalls: string[];
} {
  const dbusPayloads: string[] = [];
  const raiseCalls: string[] = [];
  const activeWindowHistory: string[] = [];
  let activeWindow: ScriptWindow | null = null;
  const workspace: ScriptWorkspace = {
    windowList: () => windows,
    raiseWindow: (window: ScriptWindow) => {
      raiseCalls.push(window.testId ?? window.caption);
    },
    windowAdded: createScriptSignal(),
    windowRemoved: createScriptSignal(),
    windowActivated: createScriptSignal(),
    screensChanged: createScriptSignal(),
  };

  Object.defineProperty(workspace, 'activeWindow', {
    configurable: true,
    enumerable: true,
    get: () => activeWindow,
    set: (window: ScriptWindow | null) => {
      activeWindow = window;
      activeWindowHistory.push(window?.testId ?? window?.caption ?? 'null');
    },
  });

  vm.runInNewContext(buildKWinBridgeScript('io.github.subminer.kwinbridge.test'), {
    Array,
    GuardedWeakSet,
    JSON,
    Number,
    Object,
    String,
    WeakSet: GuardedWeakSet,
    callDBus: (
      _serviceName: string,
      _objectPath: string,
      _interfaceName: string,
      member: string,
      payload: string,
    ) => {
      assert.equal(member, 'Update');
      dbusPayloads.push(payload);
    },
    workspace,
  });

  return { dbusPayloads, workspace, activeWindowHistory, raiseCalls };
}

function withPlatform<T>(platform: NodeJS.Platform, callback: () => T): T {
  const originalDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: platform,
  });

  try {
    return callback();
  } finally {
    if (originalDescriptor) {
      Object.defineProperty(process, 'platform', originalDescriptor);
    }
  }
}

function withEnv<T>(overrides: Record<string, string | undefined>, callback: () => T): T {
  const originalValues = new Map<string, string | undefined>();
  for (const key of Object.keys(overrides)) {
    originalValues.set(key, process.env[key]);
    const value = overrides[key];
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }

  try {
    return callback();
  } finally {
    for (const [key, value] of originalValues) {
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }
  }
}

test('selectKWinMpvWindow prefers the active window among socket matches', () => {
  const commandLines = new Map<number, string>([
    [10, 'mpv --input-ipc-server=/tmp/subminer.sock first.mkv'],
    [20, 'mpv --input-ipc-server=/tmp/subminer.sock second.mkv'],
  ]);

  const selected = selectKWinMpvWindow(
    [
      makeWindow({ pid: 10, active: false }),
      makeWindow({ pid: 20, active: true }),
    ],
    {
      targetMpvSocketPath: '/tmp/subminer.sock',
      getWindowCommandLine: (pid) => commandLines.get(pid) ?? null,
    },
  );

  assert.equal(selected?.pid, 20);
});

test('selectKWinMpvWindow requires an exact socket-path match', () => {
  const selected = selectKWinMpvWindow(
    [
      makeWindow({ pid: 10, active: true }),
      makeWindow({ pid: 20, active: false }),
    ],
    {
      targetMpvSocketPath: '/tmp/subminer.sock',
      getWindowCommandLine: (pid) => {
        if (pid === 10) return 'mpv --input-ipc-server=/tmp/subminer.sock2 second.mkv';
        if (pid === 20) return 'mpv --input-ipc-server=/tmp/subminer.sock first.mkv';
        return null;
      },
    },
  );

  assert.equal(selected?.pid, 20);
});

test('selectKWinMpvWindow matches quoted socket paths exactly', () => {
  const selected = selectKWinMpvWindow(
    [makeWindow({ pid: 10, active: true })],
    {
      targetMpvSocketPath: '/tmp/subminer socket.sock',
      getWindowCommandLine: () => 'mpv --input-ipc-server="/tmp/subminer socket.sock" first.mkv',
    },
  );

  assert.equal(selected?.pid, 10);
});

test('selectKWinMpvWindow ignores minimized and non-mpv windows', () => {
  const selected = selectKWinMpvWindow(
    [
      makeWindow({ minimized: true, pid: 1 }),
      makeWindow({ resourceClass: 'vlc', resourceName: 'vlc', caption: 'VLC media player', pid: 2 }),
      makeWindow({ pid: 3, x: 100, y: 200, width: 1920, height: 1080 }),
    ],
    {
      targetMpvSocketPath: null,
      getWindowCommandLine: () => null,
    },
  );

  assert.equal(selected?.pid, 3);
});

test('detectCompositor resolves kwin on KDE Plasma Wayland', () => {
  withPlatform('linux', () => {
    withEnv(
      {
        HYPRLAND_INSTANCE_SIGNATURE: undefined,
        SWAYSOCK: undefined,
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
        XDG_CURRENT_DESKTOP: 'KDE',
        XDG_SESSION_DESKTOP: 'KDE',
      },
      () => {
        assert.equal(detectCompositor(), 'kwin');
      },
    );
  });
});

test('detectCompositor resolves sway when SWAYSOCK is present', () => {
  withPlatform('linux', () => {
    withEnv(
      {
        HYPRLAND_INSTANCE_SIGNATURE: undefined,
        SWAYSOCK: '/tmp/sway.sock',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_SESSION_TYPE: 'wayland',
        XDG_CURRENT_DESKTOP: undefined,
        XDG_SESSION_DESKTOP: undefined,
      },
      () => {
        assert.equal(detectCompositor(), 'sway');
      },
    );
  });
});

test('KWin tracker names are instance-scoped', () => {
  assert.equal(
    buildKWinTrackerServiceName('p123_abc'),
    'io.github.subminer.kwinbridge.p123_abc',
  );
  assert.equal(
    buildKWinTrackerPluginName('p123_abc'),
    'subminerKWinTracker_p123_abc',
  );
});

test('KWin bridge script skips unsafe windows before WeakSet access', () => {
  const safeMpvWindow = makeScriptWindow({ testId: 'safe-mpv' });
  const overlayWindow = makeOverlayScriptWindow({ testId: 'overlay-window' });
  const safeCandidateWindow = makeScriptWindow({
    caption: 'Terminal',
    resourceClass: 'konsole',
    resourceName: 'konsole',
    testId: 'safe-candidate',
  });
  const deletedWindow = makeScriptWindow({
    __weakSetUnsafe: true,
    deleted: true,
    testId: 'deleted-window',
  });
  const transientWindow = makeScriptWindow({
    __weakSetUnsafe: true,
    testId: 'transient-window',
    transient: true,
  });
  const unmanagedWindow = makeScriptWindow({
    __weakSetUnsafe: true,
    managed: false,
    testId: 'unmanaged-window',
  });
  const nonNormalWindow = makeScriptWindow({
    __weakSetUnsafe: true,
    normalWindow: false,
    testId: 'non-normal-window',
  });
  const specialWindow = makeScriptWindow({
    __weakSetUnsafe: true,
    specialWindow: true,
    testId: 'special-window',
  });
  const outlineWindow = makeScriptWindow({
    __weakSetUnsafe: true,
    outline: true,
    testId: 'outline-window',
  });
  const { dbusPayloads, workspace } = runKWinBridgeScript([
    safeMpvWindow,
    overlayWindow,
    safeCandidateWindow,
    deletedWindow,
    transientWindow,
    unmanagedWindow,
    nonNormalWindow,
    specialWindow,
    outlineWindow,
  ]);

  assert.equal(workspace.windowActivated.callbacks.length > 0, true);
  assert.equal(safeMpvWindow.closed.callbacks.length > 0, true);
  assert.equal(overlayWindow.closed.callbacks.length > 0, true);
  assert.equal(safeCandidateWindow.windowClassChanged.callbacks.length, 0);
  assert.equal(deletedWindow.closed.callbacks.length, 0);
  assert.equal(transientWindow.closed.callbacks.length, 0);
  assert.equal(unmanagedWindow.closed.callbacks.length, 0);
  assert.equal(nonNormalWindow.closed.callbacks.length, 0);
  assert.equal(specialWindow.closed.callbacks.length, 0);
  assert.equal(outlineWindow.closed.callbacks.length, 0);
  assert.deepEqual(parseLastBridgePayload(dbusPayloads).windows?.map((window) => window.pid), [100]);

  const popupWindow = makeScriptWindow({
    __weakSetUnsafe: true,
    popupWindow: true,
    testId: 'popup-window',
  });
  workspace.windowAdded.emit(popupWindow);
  assert.equal(popupWindow.closed.callbacks.length, 0);
});

test('KWin bridge script watches mpv windows for active changes', () => {
  const mpvWindow = makeScriptWindow({ active: false, testId: 'mpv-window' });
  const { dbusPayloads } = runKWinBridgeScript([mpvWindow]);

  assert.equal(parseLastBridgePayload(dbusPayloads).windows?.[0]?.active, false);

  mpvWindow.active = true;
  mpvWindow.activeChanged.emit(mpvWindow);

  assert.equal(parseLastBridgePayload(dbusPayloads).windows?.[0]?.active, true);
});

test('KWin bridge script uses client geometry for overlay placement', () => {
  const mpvWindow = makeScriptWindow({
    testId: 'mpv-window',
    clientGeometry: {
      x: 40,
      y: 60,
      width: 1280,
      height: 720,
    },
    frameGeometry: {
      x: 12,
      y: 24,
      width: 1360,
      height: 816,
    },
  });
  const overlayWindow = trackWindowMutations(
    makeOverlayScriptWindow({
      testId: 'overlay-window',
      frameGeometry: {
        x: 0,
        y: 0,
        width: 10,
        height: 10,
      },
    }),
  );
  const { dbusPayloads } = runKWinBridgeScript([mpvWindow, overlayWindow]);

  assert.equal(mpvWindow.clientGeometryChanged.callbacks.length > 0, true);
  assert.deepEqual(overlayWindow.frameGeometryAssignments.at(-1), {
    x: 40,
    y: 60,
    width: 1280,
    height: 720,
  });
  assert.deepEqual(parseLastBridgePayload(dbusPayloads).windows?.[0], {
    active: false,
    caption: 'mpv',
    minimized: false,
    normalWindow: true,
    pid: 100,
    resourceClass: 'mpv',
    resourceName: 'mpv',
    x: 40,
    y: 60,
    width: 1280,
    height: 720,
  });

  mpvWindow.clientGeometry = {
    x: 55,
    y: 75,
    width: 1200,
    height: 680,
  };
  mpvWindow.clientGeometryChanged.emit(mpvWindow);

  assert.deepEqual(overlayWindow.frameGeometryAssignments.at(-1), {
    x: 55,
    y: 75,
    width: 1200,
    height: 680,
  });
  assert.deepEqual(parseLastBridgePayload(dbusPayloads).windows?.[0], {
    active: false,
    caption: 'mpv',
    minimized: false,
    normalWindow: true,
    pid: 100,
    resourceClass: 'mpv',
    resourceName: 'mpv',
    x: 55,
    y: 75,
    width: 1200,
    height: 680,
  });
});

test('KWin bridge script raises the mpv and overlay pair together when mpv is active', () => {
  const mpvWindow = makeScriptWindow({
    active: true,
    testId: 'mpv-window',
  });
  const overlayWindow = trackWindowMutations(
    makeOverlayScriptWindow({
      testId: 'overlay-window',
    }),
  );

  const { activeWindowHistory, raiseCalls } = runKWinBridgeScript([mpvWindow, overlayWindow]);

  assert.deepEqual(raiseCalls, ['mpv-window', 'overlay-window']);
  assert.deepEqual(activeWindowHistory, ['mpv-window', 'overlay-window']);
  assert.deepEqual(overlayWindow.keepAboveAssignments, [true]);
});

test('KWin bridge script hides overlay windows when mpv is minimized and restores only script-hidden overlays', () => {
  const mpvWindow = trackWindowMutations(
    makeScriptWindow({
      active: true,
      testId: 'mpv-window',
    }),
  );
  const overlayWindow = trackWindowMutations(
    makeOverlayScriptWindow({
      testId: 'overlay-window',
    }),
  );
  const manuallyHiddenOverlayWindow = trackWindowMutations(
    makeOverlayScriptWindow({
      minimized: true,
      testId: 'manual-overlay-window',
    }),
  );

  runKWinBridgeScript([mpvWindow, overlayWindow, manuallyHiddenOverlayWindow]);

  mpvWindow.minimized = true;
  mpvWindow.windowHidden.emit(mpvWindow);
  assert.equal(overlayWindow.minimizedAssignments.includes(true), true);

  mpvWindow.minimized = false;
  mpvWindow.windowShown.emit(mpvWindow);

  assert.equal(overlayWindow.minimizedAssignments.at(-1), false);
  assert.equal(manuallyHiddenOverlayWindow.minimizedAssignments.includes(false), false);
});

test('KWin bridge script restores mpv when the overlay is shown while mpv is minimized', () => {
  const mpvWindow = trackWindowMutations(
    makeScriptWindow({
      minimized: true,
      testId: 'mpv-window',
    }),
  );
  const overlayWindow = trackWindowMutations(
    makeOverlayScriptWindow({
      minimized: true,
      testId: 'overlay-window',
    }),
  );

  const { activeWindowHistory, raiseCalls } = runKWinBridgeScript([mpvWindow, overlayWindow]);

  overlayWindow.minimized = false;
  overlayWindow.windowShown.emit(overlayWindow);

  assert.equal(mpvWindow.minimizedAssignments.includes(false), true);
  assert.equal(raiseCalls.at(-2), 'mpv-window');
  assert.equal(raiseCalls.at(-1), 'overlay-window');
  assert.equal(activeWindowHistory.at(-2), 'mpv-window');
  assert.equal(activeWindowHistory.at(-1), 'overlay-window');
});

test('KWin bridge script restores overlay keep-above state after the pair loses activation', () => {
  const mpvWindow = makeScriptWindow({
    active: true,
    testId: 'mpv-window',
  });
  const overlayWindow = trackWindowMutations(
    makeOverlayScriptWindow({
      testId: 'overlay-window',
    }),
  );
  const otherWindow = makeScriptWindow({
    active: true,
    caption: 'Terminal',
    resourceClass: 'konsole',
    resourceName: 'konsole',
    testId: 'other-window',
  });

  const { workspace } = runKWinBridgeScript([mpvWindow, overlayWindow, otherWindow]);

  mpvWindow.active = false;
  overlayWindow.active = false;
  workspace.activeWindow = otherWindow;
  workspace.windowActivated.emit(otherWindow);

  assert.equal(overlayWindow.keepAboveAssignments.at(-1), false);
});

test('KWin tracker falls back to unloading unnamed loadScript calls by file path', async () => {
  const tracker = new KWinWindowTracker() as any;
  const callSignatures: string[] = [];

  tracker.callMethod = async (_bus: unknown, options: { member: string; signature?: string }) => {
    if (options.member !== 'loadScript') {
      throw new Error(`unexpected method: ${options.member}`);
    }

    callSignatures.push(options.signature ?? '');
    if (options.signature === 'ss') {
      throw new Error("Expected 1 body elements for signature 's'");
    }

    return 17;
  };

  try {
    const loadedScript = await tracker.loadScript(
      {} as never,
      '/tmp/subminer-kwin-test/main.js',
      'subminerKWinTracker_test',
    );

    assert.deepEqual(loadedScript, {
      scriptId: 17,
      unloadKey: '/tmp/subminer-kwin-test/main.js',
    });
    assert.deepEqual(callSignatures, ['ss', 's']);
  } finally {
    await tracker.stopAsync();
  }
});

test('KWin tracker rejects empty non-void DBus replies', async () => {
  const tracker = new KWinWindowTracker() as any;

  try {
    await assert.rejects(
      tracker.callMethod(
        {
          call: async () => ({
            body: [],
          }),
        } as never,
        {
          path: '/Scripting',
          interfaceName: 'org.kde.kwin.Scripting',
          member: 'loadScript',
          signature: 's',
          body: ['/tmp/subminer-kwin-test/main.js'],
        },
      ),
      /Empty reply body from org\.kde\.kwin\.Scripting\.loadScript/,
    );
  } finally {
    await tracker.stopAsync();
  }
});

test('KWin tracker stop unloads the tracked fallback script key', async () => {
  const tracker = new KWinWindowTracker() as any;
  const calls: string[] = [];

  tracker.bus = {
    disconnect: () => {
      calls.push('disconnect');
    },
    releaseName: async () => {
      calls.push('releaseName');
    },
    unexport: () => {
      calls.push('unexport');
    },
  };
  tracker.scriptId = 23;
  tracker.unloadScriptKey = '/tmp/subminer-kwin-test/main.js';
  tracker.stopScript = async (_bus: unknown, scriptId: number) => {
    calls.push(`stop:${scriptId}`);
  };
  tracker.unloadScript = async (_bus: unknown, unloadKey: string) => {
    calls.push(`unload:${unloadKey}`);
    return true;
  };

  await tracker.stopAsync();

  assert.deepEqual(calls.slice(0, 2), [
    'stop:23',
    'unload:/tmp/subminer-kwin-test/main.js',
  ]);
});

test('KWin tracker recreates its temp workspace after stop', async () => {
  const tracker = new KWinWindowTracker() as any;
  const firstScriptPath = tracker.ensureScriptWorkspace();

  assert.equal(fs.existsSync(path.dirname(firstScriptPath)), true);

  await tracker.stopAsync();

  const secondScriptPath = tracker.ensureScriptWorkspace();
  assert.notEqual(secondScriptPath, firstScriptPath);
  assert.equal(fs.existsSync(path.dirname(secondScriptPath)), true);

  await tracker.stopAsync();
});

test('KWin tracker ignores malformed windows payloads', async () => {
  const tracker = new KWinWindowTracker() as any;
  const geometries: unknown[] = [];

  tracker.updateGeometry = (geometry: unknown) => {
    geometries.push(geometry);
  };
  tracker.updateFocus = () => {};
  tracker.handleUpdate(JSON.stringify({ windows: { pid: 1 } }));

  assert.deepEqual(geometries, [null]);
  await tracker.stopAsync();
});

test('KWin tracker filters malformed window entries before selection', async () => {
  const tracker = new KWinWindowTracker() as any;
  const geometries: unknown[] = [];
  const focusStates: boolean[] = [];

  tracker.updateGeometry = (geometry: unknown) => {
    geometries.push(geometry);
  };
  tracker.updateFocus = (focused: boolean) => {
    focusStates.push(focused);
  };
  tracker.getWindowCommandLine = () => null;

  tracker.handleUpdate(
    JSON.stringify({
      windows: [null, 42, 'mpv', { pid: 7 }, makeWindow({ active: true, pid: 9, x: 50, y: 60 })],
    }),
  );

  assert.deepEqual(geometries, [{ x: 50, y: 60, width: 1280, height: 720 }]);
  assert.deepEqual(focusStates, [true]);
  await tracker.stopAsync();
});

test('KWin tracker caches process command lines between updates', async () => {
  const tracker = new KWinWindowTracker() as any;
  let readCount = 0;

  tracker.readProcessCommandLine = (pid: number) => {
    readCount += 1;
    return `mpv --input-ipc-server=/tmp/subminer-${pid}.sock`;
  };

  assert.equal(tracker.getWindowCommandLine(42), 'mpv --input-ipc-server=/tmp/subminer-42.sock');
  assert.equal(tracker.getWindowCommandLine(42), 'mpv --input-ipc-server=/tmp/subminer-42.sock');
  assert.equal(readCount, 1);

  await tracker.stopAsync();
});

test('KWin tracker rejects invalid pids before reading process command lines', async () => {
  const tracker = new KWinWindowTracker() as any;

  assert.equal(tracker.readProcessCommandLine(0), null);
  assert.equal(tracker.readProcessCommandLine(-1), null);
  assert.equal(tracker.readProcessCommandLine(1.5), null);

  await tracker.stopAsync();
});
