import { type BrowserWindow, screen } from 'electron';
import { execFile } from 'node:child_process';
import { startOverlayWindowTracker as startOverlayWindowTrackerCore } from '../../core/services';
import { isHeadlessInitialCommand, type CliArgs } from '../../cli/args';
import type { OverlayContentMeasurement, WindowGeometry } from '../../types';
import { createWindowTracker as createWindowTrackerCore } from '../../window-trackers';
import type { BaseWindowTracker } from '../../window-trackers';
import {
  bindWindowsOverlayAboveMpv,
  clearWindowsOverlayOwner,
  findWindowsMpvTargetWindowHandle,
  getWindowsForegroundProcessName,
  setWindowsOverlayOwner,
} from '../../window-trackers/windows-helper';
import {
  applyLinuxOverlayInputShape,
  applyLinuxOverlayPointerInteractionMousePassthrough,
  ensureLinuxOverlayPointerInteractionLoop,
  type ForegroundSuppressionGraceState,
  mapOverlayMeasurementForPointerInteraction,
  resolveForegroundSuppressionWithGrace,
  shouldPrimeLinuxOverlayInteractionFromMeasurement,
  tickLinuxOverlayPointerInteraction,
} from './linux-overlay-pointer-interaction';
import { restoreLinuxOverlayWindowShape } from './linux-overlay-window-shape';
import {
  ensureLinuxOverlayZOrderKeepAliveLoop,
  shouldRunLinuxOverlayZOrderKeepAlive,
  tickLinuxOverlayZOrderKeepAlive,
} from './linux-overlay-zorder-keepalive';
import { createLinuxX11CursorPointReader } from './linux-x11-cursor-point';
import type { LinuxVisibleOverlayWindowMode } from './linux-visible-overlay-window-mode';
import { createStatsOverlayVisibilityChangeHandler } from './stats-overlay-visibility';
import { hasLiveSeparateWindow } from './settings-window-z-order';
import { tickWindowsOverlayPointerInteraction } from './windows-overlay-pointer-interaction';

export interface VisibleOverlayInteractionRuntimeDeps {
  overlayManager: {
    getMainWindow: () => BrowserWindow | null;
    getVisibleOverlayVisible: () => boolean;
  };
  overlayContentMeasurementStore: {
    clear: (layer: 'visible') => void;
    getLatestByLayer: (layer: 'visible') => OverlayContentMeasurement | null;
  };
  logger: {
    info: (message: string, ...args: unknown[]) => void;
    warn: (message: string, ...args: unknown[]) => void;
    debug: (message: string, ...args: unknown[]) => void;
  };
  updateVisibleOverlayVisibility: () => void;
  getModalInputExclusive: () => boolean;
  getStatsOverlayVisible: () => boolean;
  setStatsOverlayVisible: (visible: boolean) => void;
  getWindowTracker: () => BaseWindowTracker | null;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  getMpvSocketPath: () => string;
  getBackendOverride: () => string | null;
  getInitialArgs: () => CliArgs | null;
  getOverlayRuntimeInitialized: () => boolean;
  getLinuxVisibleOverlayWindowMode: () => LinuxVisibleOverlayWindowMode;
  setLinuxVisibleOverlayOwnerBindingKey: (key: string | null) => void;
  bindVisibleOverlayToTrackedX11Window: (window: BrowserWindow) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  refreshCurrentSubtitle: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  resetLastOverlayWindowGeometry: () => void;
  enforceOverlayLayerOrder: () => void;
  getOverlayForegroundSeparateWindows: () => BrowserWindow[];
}

export function createVisibleOverlayInteractionRuntime(deps: VisibleOverlayInteractionRuntimeDeps) {
  const { overlayManager, overlayContentMeasurementStore, logger } = deps;

  const VISIBLE_OVERLAY_BLUR_REFRESH_DELAYS_MS = [0, 25, 100, 250] as const;
  const WINDOWS_VISIBLE_OVERLAY_Z_ORDER_RETRY_DELAYS_MS = [0, 48, 120, 240, 480] as const;
  const WINDOWS_VISIBLE_OVERLAY_FOREGROUND_POLL_INTERVAL_MS = 75;
  const WINDOWS_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS = 200;
  const LINUX_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS = 1_500;
  // Ignore transient "neither mpv nor overlay is the active window" blips before suppressing
  // subtitle pointer interaction. Right after playback starts the overlay can briefly become the
  // X11 active window, which would otherwise leave subtitles inert for a poll cycle (~1s).
  const LINUX_POINTER_FOREGROUND_SUPPRESS_GRACE_MS = 500;
  const LINUX_VISIBLE_OVERLAY_STARTUP_INPUT_GRACE_MS = 1_500;
  const MACOS_VISIBLE_OVERLAY_FOREGROUND_PROBE_TIMEOUT_MS = 1_200;
  let visibleOverlayBlurRefreshTimeouts: Array<ReturnType<typeof setTimeout>> = [];
  let windowsVisibleOverlayZOrderRetryTimeouts: Array<ReturnType<typeof setTimeout>> = [];
  let windowsVisibleOverlayZOrderSyncInFlight = false;
  let windowsVisibleOverlayZOrderSyncQueued = false;
  let windowsVisibleOverlayForegroundPollInterval: ReturnType<typeof setInterval> | null = null;
  let windowsOverlayPointerInteractionActive = false;
  let lastWindowsVisibleOverlayForegroundProcessName: string | null = null;
  let lastWindowsVisibleOverlayBlurredAtMs = 0;
  let lastLinuxVisibleOverlayFollowedMpvAtMs = 0;
  const linuxPointerForegroundSuppressionGrace: ForegroundSuppressionGraceState = {
    lossSinceMs: null,
  };
  let visibleOverlayInteractionActive = false;
  let linuxOverlayInputShapeActive = false;
  let linuxOverlayPointerInteractionStateApplied = process.platform !== 'linux';
  let linuxVisibleOverlayStartupInputPrimed = false;
  let linuxVisibleOverlayStartupInputGraceUntilMs = 0;
  // Renderer-reported interactive hint (Linux only): true while a Yomitan popup/modal
  // region is interactive, so the cursor poll keeps the overlay interactive even when the cursor
  // moves off measured subtitle/sidebar rects onto the popup.
  let linuxOverlayInteractiveHint = false;
  let macOSVisibleOverlayForegroundProbeActive = false;
  let macOSVisibleOverlayForegroundProbeToken = 0;
  let macOSVisibleOverlayForegroundProbeTimeout: ReturnType<typeof setTimeout> | null = null;
  const linuxVisibleOverlayOwnerBindingQueues = new WeakMap<BrowserWindow, Promise<void>>();

  const handleStatsOverlayVisibilityChanged = createStatsOverlayVisibilityChangeHandler({
    setStatsOverlayVisibleState: (visible) => {
      deps.setStatsOverlayVisible(visible);
    },
    resetVisibleOverlayInteraction: () => {
      visibleOverlayInteractionActive = false;
    },
    getMainWindow: () => overlayManager.getMainWindow(),
    updateVisibleOverlayVisibility: () => deps.updateVisibleOverlayVisibility(),
  });

  function resetVisibleOverlayInputState(): void {
    visibleOverlayInteractionActive = false;
    windowsOverlayPointerInteractionActive = false;
    linuxOverlayInputShapeActive = false;
    linuxOverlayPointerInteractionStateApplied = false;
    resetLinuxVisibleOverlayStartupInputPrimer();
    linuxOverlayInteractiveHint = false;
    overlayContentMeasurementStore.clear('visible');
    const mainWindow = overlayManager.getMainWindow();
    if (process.platform === 'linux' && mainWindow && !mainWindow.isDestroyed()) {
      restoreLinuxOverlayWindowShape(mainWindow);
    }
  }

  function restoreVisibleOverlayWindowShapeForShow(): void {
    if (process.platform !== 'linux') {
      return;
    }
    restoreLinuxOverlayWindowShape(overlayManager.getMainWindow());
  }

  function clearVisibleOverlayBlurRefreshTimeouts(): void {
    for (const timeout of visibleOverlayBlurRefreshTimeouts) {
      clearTimeout(timeout);
    }
    visibleOverlayBlurRefreshTimeouts = [];
  }

  function clearWindowsVisibleOverlayZOrderRetryTimeouts(): void {
    for (const timeout of windowsVisibleOverlayZOrderRetryTimeouts) {
      clearTimeout(timeout);
    }
    windowsVisibleOverlayZOrderRetryTimeouts = [];
  }

  function finishMacOSVisibleOverlayForegroundProbe(token: number): void {
    if (token !== macOSVisibleOverlayForegroundProbeToken) {
      return;
    }
    if (macOSVisibleOverlayForegroundProbeTimeout !== null) {
      clearTimeout(macOSVisibleOverlayForegroundProbeTimeout);
      macOSVisibleOverlayForegroundProbeTimeout = null;
    }
    if (!macOSVisibleOverlayForegroundProbeActive) {
      return;
    }
    macOSVisibleOverlayForegroundProbeActive = false;
    deps.updateVisibleOverlayVisibility();
  }

  function startMacOSVisibleOverlayForegroundProbe(): void {
    if (process.platform !== 'darwin') {
      return;
    }
    const tracker = deps.getWindowTracker();
    if (!tracker) {
      return;
    }

    macOSVisibleOverlayForegroundProbeActive = true;
    const token = ++macOSVisibleOverlayForegroundProbeToken;
    if (macOSVisibleOverlayForegroundProbeTimeout !== null) {
      clearTimeout(macOSVisibleOverlayForegroundProbeTimeout);
    }
    macOSVisibleOverlayForegroundProbeTimeout = setTimeout(() => {
      finishMacOSVisibleOverlayForegroundProbe(token);
    }, MACOS_VISIBLE_OVERLAY_FOREGROUND_PROBE_TIMEOUT_MS);

    void tracker
      .refreshNow()
      .catch((error) => {
        logger.warn('Failed to refresh macOS frontmost app after overlay blur', error);
      })
      .finally(() => {
        finishMacOSVisibleOverlayForegroundProbe(token);
      });
  }

  function getNativeWindowHandleDecimal(window: BrowserWindow): string {
    const handle = window.getNativeWindowHandle();
    return handle.length >= 8
      ? handle.readBigUInt64LE(0).toString()
      : BigInt(handle.readUInt32LE(0)).toString();
  }

  function getWindowsNativeWindowHandle(window: BrowserWindow): string {
    return getNativeWindowHandleDecimal(window);
  }

  function getWindowsNativeWindowHandleNumber(window: BrowserWindow): number {
    const handle = window.getNativeWindowHandle();
    return handle.length >= 8 ? Number(handle.readBigUInt64LE(0)) : handle.readUInt32LE(0);
  }

  function enqueueVisibleOverlayX11OwnerBindingOperation(
    window: BrowserWindow,
    args: string[],
    onError?: (error: Error) => void,
  ): void {
    const previous = linuxVisibleOverlayOwnerBindingQueues.get(window) ?? Promise.resolve();
    const operation = previous
      .catch(() => {})
      .then(
        () =>
          new Promise<void>((resolve) => {
            if (window.isDestroyed()) {
              resolve();
              return;
            }
            execFile('xprop', args, { timeout: 1500 }, (error) => {
              if (error) {
                onError?.(error);
              }
              resolve();
            });
          }),
      );
    const queued = operation.finally(() => {
      if (linuxVisibleOverlayOwnerBindingQueues.get(window) === queued) {
        linuxVisibleOverlayOwnerBindingQueues.delete(window);
      }
    });
    linuxVisibleOverlayOwnerBindingQueues.set(window, queued);
  }

  function clearVisibleOverlayX11OwnerBinding(window: BrowserWindow): void {
    if (window.isDestroyed()) return;
    enqueueVisibleOverlayX11OwnerBindingOperation(window, [
      '-id',
      getNativeWindowHandleDecimal(window),
      '-remove',
      'WM_TRANSIENT_FOR',
    ]);
  }

  function resolveWindowsOverlayBindTargetHandle(
    targetMpvSocketPath?: string | null,
  ): number | null {
    if (process.platform !== 'win32') {
      return null;
    }

    try {
      if (targetMpvSocketPath) {
        const windowTracker = deps.getWindowTracker() as {
          getTargetWindowHandle?: () => number | null;
        } | null;
        const trackedHandle = windowTracker?.getTargetWindowHandle?.();
        if (typeof trackedHandle === 'number' && Number.isFinite(trackedHandle)) {
          return trackedHandle;
        }
        return null;
      }
      return findWindowsMpvTargetWindowHandle();
    } catch {
      return null;
    }
  }

  function createOverlayWindowTracker(
    override?: string | null,
    targetMpvSocketPath?: string | null,
  ) {
    const initialArgs = deps.getInitialArgs();
    if (initialArgs && isHeadlessInitialCommand(initialArgs)) {
      return null;
    }
    return createWindowTrackerCore(override, targetMpvSocketPath);
  }

  function bindVisibleOverlayOwner(): void {
    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) return;
    if (process.platform === 'linux') {
      deps.bindVisibleOverlayToTrackedX11Window(mainWindow);
      return;
    }
    if (process.platform !== 'win32') return;
    const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
    const targetSocketPath = deps.getMpvSocketPath();
    const targetWindowHwnd = resolveWindowsOverlayBindTargetHandle(targetSocketPath);
    if (targetWindowHwnd !== null && bindWindowsOverlayAboveMpv(overlayHwnd, targetWindowHwnd)) {
      return;
    }
    if (targetSocketPath) {
      return;
    }
    const tracker = deps.getWindowTracker();
    const mpvResult = tracker
      ? (() => {
          try {
            const win32 =
              require('../../window-trackers/win32') as typeof import('../../window-trackers/win32');
            const poll = win32.findMpvWindows();
            const focused = poll.matches.find((m) => m.isForeground);
            return focused ?? [...poll.matches].sort((a, b) => b.area - a.area)[0] ?? null;
          } catch {
            return null;
          }
        })()
      : null;
    if (!mpvResult) return;
    if (!setWindowsOverlayOwner(overlayHwnd, mpvResult.hwnd)) {
      logger.warn('Failed to set overlay owner via koffi');
    }
  }

  function releaseVisibleOverlayOwner(): void {
    const mainWindow = overlayManager.getMainWindow();
    if (process.platform === 'linux') {
      deps.setLinuxVisibleOverlayOwnerBindingKey(null);
      if (mainWindow && !mainWindow.isDestroyed()) {
        clearVisibleOverlayX11OwnerBinding(mainWindow);
      }
      return;
    }
    if (process.platform !== 'win32' || !mainWindow || mainWindow.isDestroyed()) return;
    const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
    if (!clearWindowsOverlayOwner(overlayHwnd)) {
      logger.warn('Failed to clear overlay owner via koffi');
    }
  }

  function startOverlayWindowTrackerForCurrentSocket(): void {
    startOverlayWindowTrackerCore({
      backendOverride: deps.getBackendOverride(),
      getMpvSocketPath: () => deps.getMpvSocketPath(),
      createWindowTracker: createOverlayWindowTracker,
      setWindowTracker: (tracker) => {
        deps.setWindowTracker(tracker);
      },
      updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
        deps.updateVisibleOverlayBounds(geometry),
      isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      updateVisibleOverlayVisibility: () => deps.updateVisibleOverlayVisibility(),
      refreshCurrentSubtitle: () => {
        deps.refreshCurrentSubtitle();
      },
      getOverlayWindows: () => deps.getOverlayWindows(),
      syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
      bindOverlayOwner: () => bindVisibleOverlayOwner(),
      releaseOverlayOwner: () => releaseVisibleOverlayOwner(),
    });
  }

  function retargetOverlayWindowTrackerForMpvSocket(
    nextSocketPath: string,
    previousSocketPath: string,
  ): void {
    if (nextSocketPath === previousSocketPath || !deps.getOverlayRuntimeInitialized()) {
      return;
    }

    const previousTracker = deps.getWindowTracker();
    if (previousTracker) {
      try {
        previousTracker.stop();
      } catch (error) {
        logger.warn('Failed to stop previous overlay window tracker before retargeting', error);
      }
    }

    releaseVisibleOverlayOwner();
    deps.setWindowTracker(null);
    deps.setTrackerNotReadyWarningShown(false);
    deps.resetLastOverlayWindowGeometry();
    startOverlayWindowTrackerForCurrentSocket();
    deps.updateVisibleOverlayVisibility();
    deps.syncOverlayShortcuts();
    logger.info(
      `Retargeted overlay window tracker for MPV socket: ${previousSocketPath} -> ${nextSocketPath}`,
    );
  }

  async function syncWindowsVisibleOverlayToMpvZOrder(): Promise<boolean> {
    if (process.platform !== 'win32') {
      return false;
    }

    const mainWindow = overlayManager.getMainWindow();
    if (
      !mainWindow ||
      mainWindow.isDestroyed() ||
      !mainWindow.isVisible() ||
      !overlayManager.getVisibleOverlayVisible()
    ) {
      return false;
    }

    const windowTracker = deps.getWindowTracker();
    if (!windowTracker) {
      return false;
    }

    if (
      typeof windowTracker.isTargetWindowMinimized === 'function' &&
      windowTracker.isTargetWindowMinimized()
    ) {
      return false;
    }

    if (!windowTracker.isTracking() && windowTracker.getGeometry() === null) {
      return false;
    }

    const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
    const targetWindowHwnd = resolveWindowsOverlayBindTargetHandle(deps.getMpvSocketPath());
    if (targetWindowHwnd !== null && bindWindowsOverlayAboveMpv(overlayHwnd, targetWindowHwnd)) {
      (mainWindow as BrowserWindow & { setOpacity?: (opacity: number) => void }).setOpacity?.(1);
      return true;
    }
    return false;
  }

  function requestWindowsVisibleOverlayZOrderSync(): void {
    if (process.platform !== 'win32') {
      return;
    }

    if (windowsVisibleOverlayZOrderSyncInFlight) {
      windowsVisibleOverlayZOrderSyncQueued = true;
      return;
    }

    windowsVisibleOverlayZOrderSyncInFlight = true;
    void syncWindowsVisibleOverlayToMpvZOrder()
      .catch((error) => {
        logger.warn('Failed to bind Windows overlay z-order to mpv', error);
      })
      .finally(() => {
        windowsVisibleOverlayZOrderSyncInFlight = false;
        if (!windowsVisibleOverlayZOrderSyncQueued) {
          return;
        }

        windowsVisibleOverlayZOrderSyncQueued = false;
        requestWindowsVisibleOverlayZOrderSync();
      });
  }

  function scheduleWindowsVisibleOverlayZOrderSyncBurst(): void {
    if (process.platform !== 'win32') {
      return;
    }

    clearWindowsVisibleOverlayZOrderRetryTimeouts();
    for (const delayMs of WINDOWS_VISIBLE_OVERLAY_Z_ORDER_RETRY_DELAYS_MS) {
      const retryTimeout = setTimeout(() => {
        windowsVisibleOverlayZOrderRetryTimeouts = windowsVisibleOverlayZOrderRetryTimeouts.filter(
          (timeout) => timeout !== retryTimeout,
        );
        requestWindowsVisibleOverlayZOrderSync();
      }, delayMs);
      windowsVisibleOverlayZOrderRetryTimeouts.push(retryTimeout);
    }
  }

  function hasWindowsVisibleOverlayFocusHandoffGrace(): boolean {
    return (
      process.platform === 'win32' &&
      lastWindowsVisibleOverlayBlurredAtMs > 0 &&
      Date.now() - lastWindowsVisibleOverlayBlurredAtMs <=
        WINDOWS_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS
    );
  }

  function shouldPollWindowsVisibleOverlayForegroundProcess(): boolean {
    if (process.platform !== 'win32' || !overlayManager.getVisibleOverlayVisible()) {
      return false;
    }

    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) {
      return false;
    }

    const windowTracker = deps.getWindowTracker();
    if (!windowTracker) {
      return false;
    }

    if (
      typeof windowTracker.isTargetWindowMinimized === 'function' &&
      windowTracker.isTargetWindowMinimized()
    ) {
      return false;
    }

    const overlayFocused = mainWindow.isFocused();
    const trackerFocused = windowTracker.isTargetWindowFocused?.() ?? false;
    return !overlayFocused && !trackerFocused;
  }

  function maybePollWindowsVisibleOverlayForegroundProcess(): void {
    if (!shouldPollWindowsVisibleOverlayForegroundProcess()) {
      lastWindowsVisibleOverlayForegroundProcessName = null;
      return;
    }

    const processName = getWindowsForegroundProcessName();
    const normalizedProcessName = processName?.trim().toLowerCase() ?? null;
    const previousProcessName = lastWindowsVisibleOverlayForegroundProcessName;
    lastWindowsVisibleOverlayForegroundProcessName = normalizedProcessName;

    if (normalizedProcessName !== previousProcessName) {
      deps.updateVisibleOverlayVisibility();
    }
    if (normalizedProcessName === 'mpv' && previousProcessName !== 'mpv') {
      requestWindowsVisibleOverlayZOrderSync();
    }
  }

  function ensureWindowsVisibleOverlayForegroundPollLoop(): void {
    if (process.platform !== 'win32' || windowsVisibleOverlayForegroundPollInterval !== null) {
      return;
    }

    windowsVisibleOverlayForegroundPollInterval = setInterval(() => {
      maybePollWindowsVisibleOverlayForegroundProcess();
      tickWindowsOverlayPointerInteractionNow();
    }, WINDOWS_VISIBLE_OVERLAY_FOREGROUND_POLL_INTERVAL_MS);
  }

  function clearWindowsVisibleOverlayForegroundPollLoop(): void {
    if (windowsVisibleOverlayForegroundPollInterval === null) {
      return;
    }

    clearInterval(windowsVisibleOverlayForegroundPollInterval);
    windowsVisibleOverlayForegroundPollInterval = null;
  }

  function scheduleVisibleOverlayBlurRefresh(): void {
    if (process.platform !== 'win32' && process.platform !== 'darwin') {
      return;
    }

    if (process.platform === 'win32') {
      lastWindowsVisibleOverlayBlurredAtMs = Date.now();
    }
    startMacOSVisibleOverlayForegroundProbe();
    clearVisibleOverlayBlurRefreshTimeouts();
    for (const delayMs of VISIBLE_OVERLAY_BLUR_REFRESH_DELAYS_MS) {
      const refreshTimeout = setTimeout(() => {
        visibleOverlayBlurRefreshTimeouts = visibleOverlayBlurRefreshTimeouts.filter(
          (timeout) => timeout !== refreshTimeout,
        );
        deps.updateVisibleOverlayVisibility();
      }, delayMs);
      visibleOverlayBlurRefreshTimeouts.push(refreshTimeout);
    }
  }

  function shouldSuspendWindowsOverlayPointerInteraction(): boolean {
    return (
      deps.getModalInputExclusive() ||
      deps.getStatsOverlayVisible() ||
      hasLiveSeparateWindow(deps.getOverlayForegroundSeparateWindows())
    );
  }

  function updateWindowsOverlayPointerInteractionActive(active: boolean): void {
    windowsOverlayPointerInteractionActive = active;
    visibleOverlayInteractionActive = active;

    const mainWindow = overlayManager.getMainWindow();
    if (
      process.platform !== 'win32' ||
      !mainWindow ||
      mainWindow.isDestroyed() ||
      !mainWindow.isVisible()
    ) {
      deps.updateVisibleOverlayVisibility();
      return;
    }

    if (active) {
      mainWindow.setIgnoreMouseEvents(false);
    } else {
      mainWindow.setIgnoreMouseEvents(true, { forward: true });
    }
  }

  const windowsOverlayPointerInteractionDeps = {
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getMainWindow: () => overlayManager.getMainWindow(),
    getCursorScreenPoint: () => screen.getCursorScreenPoint(),
    getSubtitleMeasurement: () => overlayContentMeasurementStore.getLatestByLayer('visible'),
    shouldSuspend: shouldSuspendWindowsOverlayPointerInteraction,
    getInteractionActive: () => windowsOverlayPointerInteractionActive,
    setInteractionActive: updateWindowsOverlayPointerInteractionActive,
  };

  function tickWindowsOverlayPointerInteractionNow(): void {
    if (process.platform !== 'win32') {
      return;
    }
    if (!windowsOverlayPointerInteractionActive && visibleOverlayInteractionActive) {
      return;
    }
    tickWindowsOverlayPointerInteraction(windowsOverlayPointerInteractionDeps);
  }

  ensureWindowsVisibleOverlayForegroundPollLoop();

  const linuxX11CursorPointReader = createLinuxX11CursorPointReader();

  function getLinuxOverlayPointerMeasurement() {
    const measurement = overlayContentMeasurementStore.getLatestByLayer('visible');
    return mapOverlayMeasurementForPointerInteraction(measurement);
  }

  function shouldSuspendLinuxOverlayPointerInteraction(): boolean {
    return deps.getModalInputExclusive() || deps.getStatsOverlayVisible();
  }

  function shouldSuppressLinuxOverlayPointerInteraction(): boolean {
    return resolveForegroundSuppressionWithGrace({
      hasForegroundSeparateWindow: hasLiveSeparateWindow(
        deps.getOverlayForegroundSeparateWindows(),
      ),
      isTrackingMpvWindow: Boolean(deps.getWindowTracker()?.isTracking()),
      isMpvWindowFocused: deps.getWindowTracker()?.isTargetWindowFocused?.() !== false,
      isOverlayWindowFocused: overlayManager.getMainWindow()?.isFocused() === true,
      nowMs: Date.now(),
      graceMs: LINUX_POINTER_FOREGROUND_SUPPRESS_GRACE_MS,
      state: linuxPointerForegroundSuppressionGrace,
    });
  }

  function shouldUseLinuxOverlayInputShape(): boolean {
    // Electron's setShape is a *bounding* shape: outside the given rects no pixels are drawn, so
    // it clips the visible subtitle (and makes a dragged subtitle vanish behind the shaped
    // region). There is no input-only region API on Linux, so selective hit-testing is handled by
    // the main-process cursor poll instead. Keep this off to avoid clipping the overlay.
    return false;
  }

  function hasLinuxVisibleOverlayStartupInputGrace(): boolean {
    return (
      process.platform === 'linux' &&
      linuxVisibleOverlayStartupInputGraceUntilMs > 0 &&
      Date.now() < linuxVisibleOverlayStartupInputGraceUntilMs
    );
  }

  function clearLinuxVisibleOverlayStartupInputGrace(): void {
    linuxVisibleOverlayStartupInputGraceUntilMs = 0;
  }

  function startLinuxVisibleOverlayStartupInputGrace(): void {
    if (process.platform !== 'linux') {
      return;
    }
    linuxVisibleOverlayStartupInputGraceUntilMs =
      Date.now() + LINUX_VISIBLE_OVERLAY_STARTUP_INPUT_GRACE_MS;
    linuxOverlayPointerInteractionStateApplied = false;
  }

  function resetLinuxVisibleOverlayStartupInputPrimer(): void {
    linuxVisibleOverlayStartupInputPrimed = false;
    clearLinuxVisibleOverlayStartupInputGrace();
    if (process.platform === 'linux') {
      visibleOverlayInteractionActive = false;
      linuxOverlayInteractiveHint = false;
      linuxOverlayPointerInteractionStateApplied = false;
    }
  }

  function applyLinuxOverlayInputShapeFromLatestMeasurement(): boolean {
    if (!shouldUseLinuxOverlayInputShape()) {
      linuxOverlayInputShapeActive = false;
      return false;
    }

    const result = applyLinuxOverlayInputShape({
      getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      getMainWindow: () => overlayManager.getMainWindow(),
      getSubtitleMeasurement: getLinuxOverlayPointerMeasurement,
      getRendererInteractiveHint: () => linuxOverlayInteractiveHint,
      shouldSuspend: shouldSuspendLinuxOverlayPointerInteraction,
      shouldSuppressInteraction: shouldSuppressLinuxOverlayPointerInteraction,
    });
    linuxOverlayInputShapeActive = result.active;
    linuxOverlayPointerInteractionStateApplied = result.handled;
    return result.handled;
  }

  function updateLinuxOverlayPointerInteractionActive(active: boolean): void {
    visibleOverlayInteractionActive = active;
    if (
      process.platform === 'linux' &&
      applyLinuxOverlayPointerInteractionMousePassthrough({
        active,
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        getMainWindow: () => overlayManager.getMainWindow(),
        shouldSuspend: shouldSuspendLinuxOverlayPointerInteraction,
        shouldSuppressInteraction: shouldSuppressLinuxOverlayPointerInteraction,
        updateVisibleOverlayVisibility: () => deps.updateVisibleOverlayVisibility(),
      })
    ) {
      linuxOverlayPointerInteractionStateApplied = true;
      return;
    }

    linuxOverlayPointerInteractionStateApplied = true;
    deps.updateVisibleOverlayVisibility();
  }

  function primeLinuxOverlayPointerInteractionAfterFirstMeasurement(): void {
    if (process.platform !== 'linux') return;
    if (linuxVisibleOverlayStartupInputPrimed) return;
    if (shouldUseLinuxOverlayInputShape()) return;
    if (
      !shouldPrimeLinuxOverlayInteractionFromMeasurement({
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        getMainWindow: () => overlayManager.getMainWindow(),
        getSubtitleMeasurement: getLinuxOverlayPointerMeasurement,
        shouldSuspend: shouldSuspendLinuxOverlayPointerInteraction,
        shouldSuppressInteraction: shouldSuppressLinuxOverlayPointerInteraction,
      })
    ) {
      return;
    }

    linuxVisibleOverlayStartupInputPrimed = true;
    linuxVisibleOverlayStartupInputGraceUntilMs =
      Date.now() + LINUX_VISIBLE_OVERLAY_STARTUP_INPUT_GRACE_MS;
    updateLinuxOverlayPointerInteractionActive(true);
  }

  const linuxOverlayZOrderKeepAliveDeps = {
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getMainWindow: () => overlayManager.getMainWindow(),
    isTrackingMpvWindow: () => Boolean(deps.getWindowTracker()?.isTracking()),
    isMpvWindowFocused: () => deps.getWindowTracker()?.isTargetWindowFocused?.() !== false,
    isOverlayWindowFocused: () => overlayManager.getMainWindow()?.isFocused() === true,
    shouldSuppressReassert: () =>
      deps.getModalInputExclusive() ||
      deps.getStatsOverlayVisible() ||
      hasLiveSeparateWindow(deps.getOverlayForegroundSeparateWindows()) ||
      (visibleOverlayInteractionActive && overlayManager.getMainWindow()?.isFocused() !== true),
    raiseMpvWindow: () => {
      if (
        lastLinuxVisibleOverlayFollowedMpvAtMs > 0 &&
        Date.now() - lastLinuxVisibleOverlayFollowedMpvAtMs <=
          LINUX_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS
      ) {
        return Promise.resolve(false);
      }
      lastLinuxVisibleOverlayFollowedMpvAtMs = Date.now();
      return deps.getWindowTracker()?.raiseTargetWindow?.() ?? Promise.resolve(false);
    },
    releaseOverlayLayerOrder: () => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      mainWindow.setAlwaysOnTop(false);
      mainWindow.setFullScreen?.(false);
      mainWindow.setVisibleOnAllWorkspaces?.(false, { visibleOnFullScreen: false });
      if (
        deps.getLinuxVisibleOverlayWindowMode() === 'fullscreen-override' &&
        mainWindow.isVisible()
      ) {
        mainWindow.hide();
      }
    },
    enforceOverlayLayerOrder: () => {
      deps.enforceOverlayLayerOrder();
    },
    focusOverlayWindow: () => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed() || mainWindow.isFocused()) return;
      mainWindow.focus();
    },
  };

  function requestLinuxOverlayZOrderFollow(): void {
    if (!shouldRunLinuxOverlayZOrderKeepAlive()) return;
    void tickLinuxOverlayZOrderKeepAlive(linuxOverlayZOrderKeepAliveDeps).catch((error) => {
      logger.debug(
        'Failed to follow tracked mpv behind focused overlay:',
        error instanceof Error ? error.message : String(error),
      );
    });
  }

  ensureLinuxOverlayZOrderKeepAliveLoop(linuxOverlayZOrderKeepAliveDeps);

  const linuxOverlayPointerInteractionDeps = {
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getMainWindow: () => overlayManager.getMainWindow(),
    getCursorScreenPoint: () =>
      linuxX11CursorPointReader.getCursorScreenPoint(screen.getCursorScreenPoint()),
    getSubtitleMeasurement: getLinuxOverlayPointerMeasurement,
    getRendererInteractiveHint: () =>
      linuxOverlayInteractiveHint || hasLinuxVisibleOverlayStartupInputGrace(),
    shouldSuspend: shouldSuspendLinuxOverlayPointerInteraction,
    shouldSuppressInteraction: shouldSuppressLinuxOverlayPointerInteraction,
    shouldUseInputShape: shouldUseLinuxOverlayInputShape,
    getInteractionActive: () => visibleOverlayInteractionActive,
    isInteractionStateApplied: () => linuxOverlayPointerInteractionStateApplied,
    setInteractionActive: updateLinuxOverlayPointerInteractionActive,
  };

  function tickLinuxOverlayPointerInteractionNow(): void {
    if (applyLinuxOverlayInputShapeFromLatestMeasurement()) {
      return;
    }
    tickLinuxOverlayPointerInteraction(linuxOverlayPointerInteractionDeps);
  }

  ensureLinuxOverlayPointerInteractionLoop(linuxOverlayPointerInteractionDeps);

  return {
    handleStatsOverlayVisibilityChanged,
    resetVisibleOverlayInputState,
    restoreVisibleOverlayWindowShapeForShow,
    startMacOSVisibleOverlayForegroundProbe,
    getNativeWindowHandleDecimal,
    getWindowsNativeWindowHandle,
    getWindowsNativeWindowHandleNumber,
    enqueueVisibleOverlayX11OwnerBindingOperation,
    clearVisibleOverlayX11OwnerBinding,
    createOverlayWindowTracker,
    bindVisibleOverlayOwner,
    releaseVisibleOverlayOwner,
    startOverlayWindowTrackerForCurrentSocket,
    retargetOverlayWindowTrackerForMpvSocket,
    requestWindowsVisibleOverlayZOrderSync,
    scheduleWindowsVisibleOverlayZOrderSyncBurst,
    hasWindowsVisibleOverlayFocusHandoffGrace,
    ensureWindowsVisibleOverlayForegroundPollLoop,
    clearWindowsVisibleOverlayForegroundPollLoop,
    scheduleVisibleOverlayBlurRefresh,
    getLinuxOverlayPointerMeasurement,
    hasLinuxVisibleOverlayStartupInputGrace,
    clearLinuxVisibleOverlayStartupInputGrace,
    startLinuxVisibleOverlayStartupInputGrace,
    resetLinuxVisibleOverlayStartupInputPrimer,
    applyLinuxOverlayInputShapeFromLatestMeasurement,
    updateLinuxOverlayPointerInteractionActive,
    primeLinuxOverlayPointerInteractionAfterFirstMeasurement,
    requestLinuxOverlayZOrderFollow,
    tickWindowsOverlayPointerInteractionNow,
    tickLinuxOverlayPointerInteractionNow,
    getVisibleOverlayInteractionActive: () => visibleOverlayInteractionActive,
    setVisibleOverlayInteractionActive: (active: boolean) => {
      visibleOverlayInteractionActive = active;
      windowsOverlayPointerInteractionActive = false;
    },
    getLinuxOverlayInputShapeActive: () => linuxOverlayInputShapeActive,
    getLastWindowsVisibleOverlayForegroundProcessName: () =>
      lastWindowsVisibleOverlayForegroundProcessName,
    getMacOSVisibleOverlayForegroundProbeActive: () => macOSVisibleOverlayForegroundProbeActive,
    setLinuxOverlayInteractiveHint: (interactive: boolean) => {
      linuxOverlayInteractiveHint = interactive;
    },
  };
}

export type VisibleOverlayInteractionRuntime = ReturnType<
  typeof createVisibleOverlayInteractionRuntime
>;
