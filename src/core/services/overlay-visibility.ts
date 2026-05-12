import type { BrowserWindow } from 'electron';
import { BaseWindowTracker } from '../../window-trackers';
import { WindowGeometry } from '../../types';
import { OVERLAY_WINDOW_CONTENT_READY_FLAG } from './overlay-window-flags';

const WINDOWS_OVERLAY_REVEAL_DELAY_MS = 48;
const pendingWindowsOverlayRevealTimeoutByWindow = new WeakMap<
  BrowserWindow,
  ReturnType<typeof setTimeout>
>();
function setOverlayWindowOpacity(window: BrowserWindow, opacity: number): void {
  const opacityCapableWindow = window as BrowserWindow & {
    setOpacity?: (opacity: number) => void;
  };
  opacityCapableWindow.setOpacity?.(opacity);
}

function clearPendingWindowsOverlayReveal(window: BrowserWindow): void {
  const pendingTimeout = pendingWindowsOverlayRevealTimeoutByWindow.get(window);
  if (!pendingTimeout) {
    return;
  }
  clearTimeout(pendingTimeout);
  pendingWindowsOverlayRevealTimeoutByWindow.delete(window);
}

function scheduleWindowsOverlayReveal(
  window: BrowserWindow,
  onReveal?: (window: BrowserWindow) => void,
): void {
  clearPendingWindowsOverlayReveal(window);
  const timeout = setTimeout(() => {
    pendingWindowsOverlayRevealTimeoutByWindow.delete(window);
    if (window.isDestroyed() || !window.isVisible()) {
      return;
    }
    setOverlayWindowOpacity(window, 1);
    onReveal?.(window);
  }, WINDOWS_OVERLAY_REVEAL_DELAY_MS);
  pendingWindowsOverlayRevealTimeoutByWindow.set(window, timeout);
}

function isOverlayWindowContentReady(window: BrowserWindow): boolean {
  return (
    (window as BrowserWindow & { [OVERLAY_WINDOW_CONTENT_READY_FLAG]?: boolean })[
      OVERLAY_WINDOW_CONTENT_READY_FLAG
    ] === true
  );
}

export function updateVisibleOverlayVisibility(args: {
  visibleOverlayVisible: boolean;
  modalActive?: boolean;
  forceMousePassthrough?: boolean;
  mainWindow: BrowserWindow | null;
  windowTracker: BaseWindowTracker | null;
  lastKnownWindowsForegroundProcessName?: string | null;
  windowsOverlayProcessName?: string | null;
  windowsFocusHandoffGraceActive?: boolean;
  trackerNotReadyWarningShown: boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  syncWindowsOverlayToMpvZOrder?: (window: BrowserWindow) => void;
  syncPrimaryOverlayWindowLayer: (layer: 'visible') => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
  isMacOSPlatform?: boolean;
  isWindowsPlatform?: boolean;
  showOverlayLoadingOsd?: (message: string) => void;
  shouldShowOverlayLoadingOsd?: () => boolean;
  markOverlayLoadingOsdShown?: () => void;
  resetOverlayLoadingOsdSuppression?: () => void;
  resolveFallbackBounds?: () => WindowGeometry;
}): void {
  if (!args.mainWindow || args.mainWindow.isDestroyed()) {
    return;
  }

  const mainWindow = args.mainWindow;

  if (args.modalActive) {
    if (args.isWindowsPlatform) {
      clearPendingWindowsOverlayReveal(mainWindow);
      setOverlayWindowOpacity(mainWindow, 0);
    }
    mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  const showPassiveVisibleOverlay = (): boolean => {
    const forceMousePassthrough = args.forceMousePassthrough === true;
    const wasVisible = mainWindow.isVisible();
    const isVisibleOverlayFocused =
      typeof mainWindow.isFocused === 'function' && mainWindow.isFocused();
    const isTrackedMacOSTargetFocused =
      !args.isMacOSPlatform || !args.windowTracker
        ? true
        : (args.windowTracker.isTargetWindowFocused?.() ?? true);
    const shouldReleaseMacOSOverlayLevel =
      args.isMacOSPlatform &&
      !!args.windowTracker &&
      !isVisibleOverlayFocused &&
      !isTrackedMacOSTargetFocused;
    const shouldDefaultToPassthrough =
      args.isWindowsPlatform || forceMousePassthrough || shouldReleaseMacOSOverlayLevel;
    const windowsForegroundProcessName =
      args.lastKnownWindowsForegroundProcessName?.trim().toLowerCase() ?? null;
    const windowsOverlayProcessName = args.windowsOverlayProcessName?.trim().toLowerCase() ?? null;
    const hasWindowsForegroundProcessSignal =
      args.isWindowsPlatform && windowsForegroundProcessName !== null;
    const isTrackedWindowsTargetFocused = args.windowTracker?.isTargetWindowFocused?.() ?? true;
    const isTrackedWindowsTargetMinimized =
      args.isWindowsPlatform &&
      typeof args.windowTracker?.isTargetWindowMinimized === 'function' &&
      args.windowTracker.isTargetWindowMinimized();
    const shouldPreserveWindowsOverlayDuringFocusHandoff =
      args.isWindowsPlatform &&
      args.windowsFocusHandoffGraceActive === true &&
      !!args.windowTracker &&
      (!hasWindowsForegroundProcessSignal ||
        windowsForegroundProcessName === 'mpv' ||
        (windowsOverlayProcessName !== null &&
          windowsForegroundProcessName === windowsOverlayProcessName)) &&
      !isTrackedWindowsTargetMinimized &&
      (args.windowTracker.isTracking() || args.windowTracker.getGeometry() !== null);
    const shouldForcePassiveReshow = args.isWindowsPlatform && !wasVisible;
    const shouldIgnoreMouseEvents =
      forceMousePassthrough ||
      (shouldDefaultToPassthrough && (!isVisibleOverlayFocused || shouldForcePassiveReshow));
    const shouldBindTrackedWindowsOverlay = args.isWindowsPlatform && !!args.windowTracker;
    const shouldKeepTrackedWindowsOverlayTopmost =
      !args.isWindowsPlatform ||
      !args.windowTracker ||
      isVisibleOverlayFocused ||
      isTrackedWindowsTargetFocused ||
      shouldPreserveWindowsOverlayDuringFocusHandoff ||
      (hasWindowsForegroundProcessSignal && windowsForegroundProcessName === 'mpv');
    if (shouldIgnoreMouseEvents) {
      mainWindow.setIgnoreMouseEvents(true, { forward: true });
    } else {
      mainWindow.setIgnoreMouseEvents(false);
    }

    if (shouldBindTrackedWindowsOverlay) {
      // On Windows, z-order is enforced by the OS via the owner window mechanism
      // (SetWindowLongPtr GWLP_HWNDPARENT). The overlay is always above mpv
      // without any manual z-order management.
    } else if (!forceMousePassthrough && !shouldReleaseMacOSOverlayLevel) {
      args.ensureOverlayWindowLevel(mainWindow);
    } else {
      mainWindow.setAlwaysOnTop(false);
    }
    if (!wasVisible) {
      const hasWebContents =
        typeof (mainWindow as unknown as { webContents?: unknown }).webContents === 'object';
      if (
        args.isWindowsPlatform &&
        hasWebContents &&
        !isOverlayWindowContentReady(mainWindow as unknown as import('electron').BrowserWindow)
      ) {
        // skip — ready-to-show hasn't fired yet; the onWindowContentReady
        // callback will trigger another visibility update when the renderer
        // has painted its first frame.
      } else if (args.isWindowsPlatform && shouldIgnoreMouseEvents) {
        setOverlayWindowOpacity(mainWindow, 0);
        mainWindow.showInactive();
        mainWindow.setIgnoreMouseEvents(true, { forward: true });
        scheduleWindowsOverlayReveal(
          mainWindow,
          shouldBindTrackedWindowsOverlay
            ? (window) => args.syncWindowsOverlayToMpvZOrder?.(window)
            : undefined,
        );
      } else {
        if (args.isWindowsPlatform) {
          setOverlayWindowOpacity(mainWindow, 0);
        }
        mainWindow.show();
        if (args.isWindowsPlatform) {
          scheduleWindowsOverlayReveal(
            mainWindow,
            shouldBindTrackedWindowsOverlay
              ? (window) => args.syncWindowsOverlayToMpvZOrder?.(window)
              : undefined,
          );
        }
      }
    }

    if (shouldBindTrackedWindowsOverlay) {
      args.syncWindowsOverlayToMpvZOrder?.(mainWindow);
    }

    if (!args.isWindowsPlatform && !args.isMacOSPlatform && !forceMousePassthrough) {
      mainWindow.focus();
    }

    return !shouldReleaseMacOSOverlayLevel;
  };

  const maybeShowOverlayLoadingOsd = (): void => {
    if (!args.isMacOSPlatform || !args.showOverlayLoadingOsd) {
      return;
    }
    if (args.shouldShowOverlayLoadingOsd && !args.shouldShowOverlayLoadingOsd()) {
      return;
    }
    args.showOverlayLoadingOsd('Overlay loading...');
    args.markOverlayLoadingOsdShown?.();
  };

  if (!args.visibleOverlayVisible) {
    args.setTrackerNotReadyWarningShown(false);
    args.resetOverlayLoadingOsdSuppression?.();
    if (args.isWindowsPlatform) {
      clearPendingWindowsOverlayReveal(mainWindow);
      setOverlayWindowOpacity(mainWindow, 0);
    }
    mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  if (args.windowTracker && args.windowTracker.isTracking()) {
    if (
      args.isWindowsPlatform &&
      typeof args.windowTracker.isTargetWindowMinimized === 'function' &&
      args.windowTracker.isTargetWindowMinimized()
    ) {
      clearPendingWindowsOverlayReveal(mainWindow);
      setOverlayWindowOpacity(mainWindow, 0);
      mainWindow.hide();
      args.syncOverlayShortcuts();
      return;
    }
    args.setTrackerNotReadyWarningShown(false);
    const geometry = args.windowTracker.getGeometry();
    if (geometry) {
      args.updateVisibleOverlayBounds(geometry);
    }
    args.syncPrimaryOverlayWindowLayer('visible');
    const shouldEnforceLayerOrder = showPassiveVisibleOverlay();
    if (shouldEnforceLayerOrder && !args.forceMousePassthrough && !args.isWindowsPlatform) {
      args.enforceOverlayLayerOrder();
    }
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.windowTracker) {
    if (args.isMacOSPlatform || args.isWindowsPlatform) {
      if (!args.trackerNotReadyWarningShown) {
        args.setTrackerNotReadyWarningShown(true);
        maybeShowOverlayLoadingOsd();
      }
      if (args.isWindowsPlatform) {
        clearPendingWindowsOverlayReveal(mainWindow);
        setOverlayWindowOpacity(mainWindow, 0);
      }
      mainWindow.hide();
      args.syncOverlayShortcuts();
      return;
    }
    args.setTrackerNotReadyWarningShown(false);
    args.syncPrimaryOverlayWindowLayer('visible');
    showPassiveVisibleOverlay();
    args.enforceOverlayLayerOrder();
    args.syncOverlayShortcuts();
    return;
  }

  const hasRetainedTrackedGeometry = args.windowTracker.getGeometry() !== null;
  const hasActiveMacOSTargetSignal =
    args.isMacOSPlatform && (args.windowTracker.isTargetWindowFocused?.() ?? false);
  const shouldPreserveTransientTrackedOverlay =
    (args.isMacOSPlatform &&
      (hasRetainedTrackedGeometry || (mainWindow.isVisible() && hasActiveMacOSTargetSignal))) ||
    (args.isWindowsPlatform &&
      typeof args.windowTracker.isTargetWindowMinimized === 'function' &&
      !args.windowTracker.isTargetWindowMinimized());

  if (
    shouldPreserveTransientTrackedOverlay &&
    (mainWindow.isVisible() || hasRetainedTrackedGeometry)
  ) {
    args.setTrackerNotReadyWarningShown(false);
    const geometry = args.windowTracker.getGeometry();
    if (geometry) {
      args.updateVisibleOverlayBounds(geometry);
    }
    args.syncPrimaryOverlayWindowLayer('visible');
    const shouldEnforceLayerOrder = showPassiveVisibleOverlay();
    if (shouldEnforceLayerOrder && !args.forceMousePassthrough && !args.isWindowsPlatform) {
      args.enforceOverlayLayerOrder();
    }
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.trackerNotReadyWarningShown) {
    args.setTrackerNotReadyWarningShown(true);
    maybeShowOverlayLoadingOsd();
  }

  if (args.isWindowsPlatform) {
    clearPendingWindowsOverlayReveal(mainWindow);
    setOverlayWindowOpacity(mainWindow, 0);
  }
  mainWindow.hide();
  args.syncOverlayShortcuts();
}

export function setVisibleOverlayVisible(options: {
  visible: boolean;
  setVisibleOverlayVisibleState: (visible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
}): void {
  options.setVisibleOverlayVisibleState(options.visible);
  options.updateVisibleOverlayVisibility();
}
