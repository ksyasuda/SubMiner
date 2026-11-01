import assert from 'node:assert/strict';
import test, { type TestContext } from 'node:test';
import type { BrowserWindow } from 'electron';
import type { HyprlandPlacementStatus } from '../../core/services/hyprland-window-placement';
import type { OverlayContentMeasurement, WindowGeometry } from '../../types';
import { createAppState } from '../state';
import { createOverlayAssembly, type OverlayAssemblyDeps } from './overlay';

type Call = { name: string; args: unknown[] };
type OverlayWindowOptions = Parameters<OverlayAssemblyDeps['createOverlayWindowCore']>[1];
type Overlay = ReturnType<typeof createOverlayAssembly>;

/** Records every dep call by name, in order; each dep still returns an inert value. */
function createRecorder() {
  const calls: Call[] = [];
  const spy =
    <A extends unknown[], R>(name: string, impl: (...args: A) => R) =>
    (...args: A): R => {
      calls.push({ name, args });
      return impl(...args);
    };
  const names = () => calls.map((call) => call.name);
  return { calls, spy, names };
}

type Recorder = ReturnType<typeof createRecorder>;

/** Test-controlled state the fakes read from. */
type FixtureState = {
  visible: boolean;
  trackedGeometry: WindowGeometry | null;
  /** Window returned by window creation; null makes creation throw after recording options. */
  window: BrowserWindow | null;
  /** The overlay manager's current main window. */
  mainWindow: BrowserWindow | null;
  /** Whether the measurement store accepts a reported content measurement. */
  measurementAccepted: boolean;
  /** Whether the plugin panel recognizes a closed panel id as the one showing. */
  pluginPanelCurrent: boolean;
};

const ACCEPTED_MEASUREMENT = {} as OverlayContentMeasurement;
const WINDOW_CREATION_STOPPED = 'window creation stopped by test';
const NO_PLACEMENT: HyprlandPlacementStatus = {
  applicable: false,
  clientFound: false,
  dispatched: false,
};

function makeDeps(recorder: Recorder, state: FixtureState) {
  const { spy } = recorder;
  const noop = () => {};
  const resolved = async () => {};
  const windowOptions: OverlayWindowOptions[] = [];
  const deps: OverlayAssemblyDeps = {
    appState: createAppState({
      mpvSocketPath: '/tmp/subminer-overlay-test.sock',
      texthookerPort: 0,
    }),
    logger: {
      debug: spy('logger.debug', noop),
      info: spy('logger.info', noop),
      warn: spy('logger.warn', noop),
    },
    overlayManager: {
      getMainWindow: spy('overlay.getMainWindow', () => state.mainWindow),
      setMainWindow: spy('overlay.setMainWindow', noop),
      getModalWindow: spy('overlay.getModalWindow', () => null),
      setModalWindow: spy('overlay.setModalWindow', noop),
      getVisibleOverlayVisible: spy('overlay.getVisibleOverlayVisible', () => state.visible),
      setVisibleOverlayVisible: spy('overlay.setVisibleOverlayVisible', (visible: boolean) => {
        state.visible = visible;
      }),
      setOverlayWindowBounds: spy('overlay.setOverlayWindowBounds', () => NO_PLACEMENT),
      setModalWindowBounds: spy('overlay.setModalWindowBounds', () => NO_PLACEMENT),
    },
    overlayModalInputState: { getModalInputExclusive: spy('modal.getInputExclusive', () => false) },
    overlayModalRuntime: {
      notifyOverlayModalOpened: spy('modal.notifyOpened', noop),
      handleOverlayModalClosed: spy('modal.handleClosed', noop),
    },
    overlayContentMeasurementStore: {
      report: spy('measurement.report', () =>
        state.measurementAccepted ? ACCEPTED_MEASUREMENT : null,
      ),
    },
    autoplayReadyGate: {
      markCurrentMediaAutoplayReady: spy('autoplay.markCurrentMediaReady', noop),
      flushPendingAutoplayReadySignal: spy('autoplay.flushPending', noop),
    },
    screen: {
      getCursorScreenPoint: spy('screen.getCursorScreenPoint', () => ({ x: 0, y: 0 })),
      getDisplayNearestPoint: spy('screen.getDisplayNearestPoint', (): never => {
        throw new Error('display fallback used');
      }),
    },
    isDev: false,
    createOverlayWindowCore: spy('createOverlayWindowCore', (_kind, options): BrowserWindow => {
      windowOptions.push(options);
      if (!state.window) throw new Error(WINDOW_CREATION_STOPPED);
      return state.window;
    }),
    getYomitanSession: spy('getYomitanSession', () => null),
    getVisibleOverlayInteraction: spy('getVisibleOverlayInteraction', () => ({
      getLinuxOverlayInputShapeActive: spy('interaction.getLinuxInputShapeActive', () => false),
      getMacOSVisibleOverlayForegroundProbeActive: spy('interaction.getMacProbe', () => false),
      getVisibleOverlayInteractionActive: spy('interaction.getActive', () => false),
      setVisibleOverlayInteractionActive: spy('interaction.setActive', noop),
      setLinuxOverlayInteractiveHint: spy('interaction.setLinuxInteractiveHint', noop),
      applyLinuxOverlayInputShapeFromLatestMeasurement: spy(
        'interaction.applyLinuxShape',
        () => false,
      ),
      tickLinuxOverlayPointerInteractionNow: spy('interaction.tickLinuxNow', noop),
      tickWindowsOverlayPointerInteractionNow: spy('interaction.tickWindowsNow', noop),
      getWindowsNativeWindowHandleNumber: spy('interaction.getWindowsHandle', () => 0),
      primeLinuxOverlayPointerInteractionAfterFirstMeasurement: spy(
        'interaction.primePointerAfterFirstMeasurement',
        noop,
      ),
      requestLinuxOverlayZOrderFollow: spy('interaction.requestLinuxZOrderFollow', noop),
      requestWindowsVisibleOverlayZOrderSync: spy('interaction.requestWindowsZOrderSync', noop),
      resetLinuxVisibleOverlayStartupInputPrimer: spy('interaction.resetStartupPrimer', noop),
      resetVisibleOverlayInputState: spy('interaction.resetInputState', noop),
      restoreVisibleOverlayWindowShapeForShow: spy('interaction.restoreShapeForShow', noop),
      scheduleVisibleOverlayBlurRefresh: spy('interaction.scheduleBlurRefresh', noop),
      startLinuxVisibleOverlayStartupInputGrace: spy('interaction.startStartupGrace', noop),
    })),
    getOverlayGeometry: spy('getOverlayGeometry', () => ({
      enforceOverlayLayerOrder: spy('geometry.enforceLayerOrder', noop),
      ensureOverlayWindowLevel: spy('geometry.ensureWindowLevel', noop),
      getCurrentOverlayGeometry: spy('geometry.getCurrentOverlayGeometry', () => ({
        x: 0,
        y: 0,
        width: 1920,
        height: 1080,
      })),
      getCurrentTrackedOverlayGeometry: spy(
        'geometry.getCurrentTrackedOverlayGeometry',
        () => state.trackedGeometry,
      ),
      syncPrimaryOverlayWindowLayer: spy('geometry.syncPrimaryLayer', noop),
      updateVisibleOverlayBounds: spy('geometry.updateVisibleBounds', () => 'not-applicable'),
    })),
    getOverlayShortcuts: spy('getOverlayShortcuts', () => ({
      syncOverlayShortcuts: spy('shortcuts.sync', noop),
      tryHandleOverlayShortcutLocalFallback: spy('shortcuts.localFallback', () => false),
    })),
    getOverlayNotifications: spy('getOverlayNotifications', () => ({
      showOverlayLoadingStatusNotification: spy('notifications.showLoading', noop),
      dismissOverlayLoadingStatusNotification: spy('notifications.dismissLoading', noop),
      flushQueuedOverlayNotifications: spy('notifications.flushQueued', noop),
      maybeStartOverlayLoadingOsd: spy('notifications.maybeStartLoadingOsd', noop),
    })),
    getOverlayVisibilityComposer: spy('getOverlayVisibilityComposer', () => ({
      broadcastRuntimeOptionsChanged: spy('composer.broadcastRuntimeOptions', noop),
      setOverlayDebugVisualizationEnabled: spy('composer.setDebugVisualization', noop),
    })),
    getMediaTimingReviewRuntime: spy('getMediaTimingReviewRuntime', () => ({
      dispose: spy('mediaTimingReview.dispose', resolved),
    })),
    getSubtitleProcessingController: spy('getSubtitleProcessingController', () => ({
      refreshCurrentSubtitle: spy('subtitle.refreshCurrent', () => true),
    })),
    getAutoplaySubtitlePriming: spy('getAutoplaySubtitlePriming', () => ({
      cancelVisibleOverlaySubtitleRefreshAfterFirstPaint: spy('priming.cancelFirstPaint', noop),
      primeCurrentSubtitleForVisibleOverlay: spy('priming.prime', resolved),
      scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint: spy('priming.scheduleFirstPaint', noop),
    })),
    getPluginPanel: spy('getPluginPanel', () => ({
      handleModalClosed: spy('pluginPanel.handleModalClosed', () => state.pluginPanelCurrent),
    })),
    setSubtitleSidebarRequestedOpen: spy('setSubtitleSidebarRequestedOpen', noop),
    stealAppFocus: spy('stealAppFocus', noop),
    ensureOverlayMpvSubtitlesHidden: spy('ensureOverlayMpvSubtitlesHidden', resolved),
    syncOverlayMpvSubtitleSuppression: spy('syncOverlayMpvSubtitleSuppression', noop),
    initializeOverlayRuntime: spy('initializeOverlayRuntime', noop),
  };
  return { deps, windowOptions };
}

/** Builds the assembly, optionally with `process.platform` overridden for the test. */
function setup(t: TestContext, options: { platform?: NodeJS.Platform; visible?: boolean } = {}) {
  if (options.platform) {
    const originalPlatform = process.platform;
    Object.defineProperty(process, 'platform', { value: options.platform, configurable: true });
    t.after(() => Object.defineProperty(process, 'platform', { value: originalPlatform }));
  }
  const recorder = createRecorder();
  const state: FixtureState = {
    visible: options.visible ?? false,
    trackedGeometry: null,
    window: null,
    mainWindow: null,
    measurementAccepted: true,
    pluginPanelCurrent: true,
  };
  const { deps, windowOptions } = makeDeps(recorder, state);
  const overlay = createOverlayAssembly(deps);
  return { recorder, state, overlay, windowOptions };
}

/** Asserts each name was called, in this relative order. */
function assertCalledInOrder(names: string[], expected: string[]): void {
  let previous = -1;
  for (const name of expected) {
    const index = names.indexOf(name, previous + 1);
    assert.ok(index > previous, `${name} was not called after the previous step`);
    previous = index;
  }
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createOverlayAssembly does not call its deps while constructing', (t) => {
  const { recorder, overlay } = setup(t);
  assert.deepEqual(recorder.names(), []);
  assert.equal(typeof overlay.setVisibleOverlayVisible, 'function');
});

const hidePaths: Array<{ label: string; hide: (overlay: Overlay) => void; order: string[] }> = [
  {
    label: 'setVisibleOverlayVisible(false)',
    hide: (overlay) => overlay.setVisibleOverlayVisible(false),
    order: [
      'autoplay.markCurrentMediaReady',
      'priming.cancelFirstPaint',
      'interaction.resetInputState',
    ],
  },
  {
    label: 'toggleVisibleOverlay() while shown',
    hide: (overlay) => overlay.toggleVisibleOverlay(),
    order: [
      'autoplay.markCurrentMediaReady',
      'priming.cancelFirstPaint',
      'interaction.resetInputState',
    ],
  },
  {
    label: 'setOverlayVisible(false)',
    hide: (overlay) => overlay.setOverlayVisible(false),
    order: [
      'priming.cancelFirstPaint',
      'interaction.resetInputState',
      'autoplay.markCurrentMediaReady',
    ],
  },
];

for (const { label, hide, order } of hidePaths) {
  test(`${label} dismisses the loading OSD, releases autoplay and clears input state`, (t) => {
    const { recorder, overlay } = setup(t, { visible: true });

    hide(overlay);

    const names = recorder.names();
    assertCalledInOrder(names, ['notifications.dismissLoading', ...order]);
    assert.ok(!names.includes('priming.prime'));
  });

  test(`${label} cancels a pending Linux fullscreen refresh burst`, async (t) => {
    const { recorder, overlay } = setup(t, { platform: 'linux', visible: true });
    overlay.handleMpvFullscreenChange(true);

    hide(overlay);
    const afterHide = recorder.calls.length;
    await sleep(20);

    // A burst that survived the hide would re-sync the overlay on its 0ms tick.
    assert.equal(recorder.calls.length, afterHide);
  });
}

// Control for the cancellation tests above: the same tick does run when nothing cancels it.
test('an uncancelled Linux fullscreen refresh burst runs on its first tick', async (t) => {
  const { recorder, overlay } = setup(t, { platform: 'linux' });
  t.after(() => overlay.clearLinuxFullscreenOverlayRefresh());
  overlay.handleMpvFullscreenChange(true);
  const afterSchedule = recorder.calls.length;

  await sleep(20);

  assert.ok(recorder.calls.length > afterSchedule);
});

const showPaths: Array<{ label: string; show: (overlay: Overlay) => void }> = [
  {
    label: 'setVisibleOverlayVisible(true)',
    show: (overlay) => overlay.setVisibleOverlayVisible(true),
  },
  {
    label: 'toggleVisibleOverlay() while hidden',
    show: (overlay) => overlay.toggleVisibleOverlay(),
  },
  { label: 'setOverlayVisible(true)', show: (overlay) => overlay.setOverlayVisible(true) },
];

for (const { label, show } of showPaths) {
  test(`${label} starts the loading OSD and input grace before priming, keeping autoplay held`, (t) => {
    const { recorder, overlay } = setup(t);

    show(overlay);

    const names = recorder.names();
    assertCalledInOrder(names, [
      'notifications.maybeStartLoadingOsd',
      'interaction.resetStartupPrimer',
      'interaction.startStartupGrace',
      'priming.prime',
    ]);
    assert.ok(!names.includes('autoplay.markCurrentMediaReady'));
    assert.ok(!names.includes('interaction.resetInputState'));
  });
}

test('visible overlay content-ready flushes notifications and autoplay without tokenizing', (t) => {
  const { recorder, overlay, windowOptions } = setup(t, { visible: true });
  assert.throws(() => overlay.createMainWindow(), new RegExp(WINDOW_CREATION_STOPPED));
  const onContentReady = windowOptions[0]?.onWindowContentReady;
  assert.ok(onContentReady);
  const before = recorder.calls.length;

  onContentReady();

  const names = recorder.names().slice(before);
  assertCalledInOrder(names, [
    'notifications.dismissLoading',
    'notifications.flushQueued',
    'interaction.primePointerAfterFirstMeasurement',
    'autoplay.flushPending',
  ]);
  // The first subtitle waits for the first content measurement, not for content-ready.
  assert.ok(!names.includes('subtitle.refreshCurrent'));
});

for (const trackedGeometry of [null, { x: 10, y: 20, width: 1280, height: 720 }]) {
  test(`Linux fullscreen recreation resets input first and ${trackedGeometry ? 'applies tracked mpv bounds' : 'keeps bounds without tracked geometry'}`, (t) => {
    const { recorder, state, overlay } = setup(t, { platform: 'linux', visible: true });
    state.trackedGeometry = trackedGeometry;
    // An opaque token: the assembly only hands the created window to the overlay manager.
    state.window = {} as BrowserWindow;

    overlay.linuxOverlayModeRuntime.sync(true);

    const names = recorder.names();
    assertCalledInOrder(names, [
      'interaction.resetInputState',
      'createOverlayWindowCore',
      'geometry.getCurrentTrackedOverlayGeometry',
      'ensureOverlayMpvSubtitlesHidden',
    ]);
    const boundsCalls = recorder.calls.filter(
      (call) => call.name === 'overlay.setOverlayWindowBounds',
    );
    assert.deepEqual(
      boundsCalls.map((call) => call.args[0]),
      trackedGeometry ? [trackedGeometry] : [],
    );
    // Recreation must never fall back to display bounds before mpv geometry is known.
    assert.ok(!names.includes('geometry.getCurrentOverlayGeometry'));
  });
}

test('an accepted content measurement refreshes pointer input, then releases autoplay, then schedules the first-paint refresh', (t) => {
  const { recorder, overlay } = setup(t);

  overlay.ipc.reportOverlayContentBounds({ layer: 'visible' });

  assertCalledInOrder(recorder.names(), [
    'measurement.report',
    'interaction.tickLinuxNow',
    'interaction.tickWindowsNow',
    'interaction.primePointerAfterFirstMeasurement',
    'autoplay.flushPending',
    'priming.scheduleFirstPaint',
  ]);
});

test('a rejected content measurement changes nothing', (t) => {
  const { recorder, state, overlay } = setup(t);
  state.measurementAccepted = false;

  overlay.ipc.reportOverlayContentBounds({ layer: 'visible' });

  assert.deepEqual(recorder.names(), ['measurement.report']);
});

test('only the main overlay window opens and closes the subtitle sidebar state', (t) => {
  const { recorder, state, overlay } = setup(t);
  const mainWindow = { isDestroyed: () => false } as BrowserWindow;
  const otherWindow = { isDestroyed: () => false } as BrowserWindow;
  state.mainWindow = mainWindow;
  const sidebarStates = () =>
    recorder.calls
      .filter((call) => call.name === 'setSubtitleSidebarRequestedOpen')
      .map((call) => call.args[0]);

  overlay.ipc.onOverlayModalOpened('subtitle-sidebar', otherWindow);
  overlay.ipc.onOverlayModalOpened('subtitle-sidebar', mainWindow);
  overlay.ipc.onOverlayModalClosed('subtitle-sidebar', otherWindow);
  overlay.ipc.onOverlayModalClosed('subtitle-sidebar', mainWindow);

  assert.deepEqual(sidebarStates(), [true, false]);
});

test('a late close for a replaced plugin panel leaves the showing panel open', (t) => {
  const { recorder, state, overlay } = setup(t);
  state.pluginPanelCurrent = false;

  overlay.ipc.onOverlayModalClosed('plugin-panel', null, 7);

  assert.equal(recorder.names().includes('modal.handleClosed'), false);
});
