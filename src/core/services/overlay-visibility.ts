import type { BrowserWindow } from 'electron';
import { BaseWindowTracker } from '../../window-trackers';
import { WindowGeometry } from '../../types';
import { OVERLAY_WINDOW_CONTENT_READY_FLAG } from './overlay-window-flags';

const WINDOWS_OVERLAY_REVEAL_DELAY_MS = 48;
const pendingWindowsOverlayRevealTimeoutByWindow = new WeakMap<
  BrowserWindow,
  ReturnType<typeof setTimeout>
>();
const pendingFirstShowBoundsRefreshGeometry = new WeakMap<BrowserWindow, WindowGeometry>();
function setOverlayWindowOpacity(window: BrowserWindow, opacity: number): void {
  const opacityCapableWindow = window as BrowserWindow & {
    setOpacity?: (opacity: number) => void;
  };
  opacityCapableWindow.setOpacity?.(opacity);
}

function releaseOverlayWindowLevel(window: BrowserWindow): void {
  window.setAlwaysOnTop(false);
  const fullscreenWindow = window as BrowserWindow & {
    setFullScreen?: (fullscreen: boolean) => void;
  };
  fullscreenWindow.setFullScreen?.(false);
  const allWorkspacesWindow = window as BrowserWindow & {
    setVisibleOnAllWorkspaces?: (
      visible: boolean,
      options?: { visibleOnFullScreen?: boolean },
    ) => void;
  };
  allWorkspacesWindow.setVisibleOnAllWorkspaces?.(false, { visibleOnFullScreen: false });
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
  nonNativeInputRegionActive?: boolean;
  suspendVisibleOverlay?: boolean;
  overlayInteractionActive?: boolean;
  mainWindow: BrowserWindow | null;
  windowTracker: BaseWindowTracker | null;
  lastKnownWindowsForegroundProcessName?: string | null;
  windowsOverlayProcessName?: string | null;
  windowsFocusHandoffGraceActive?: boolean;
  macOSForegroundProbeActive?: boolean;
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
  dismissOverlayLoadingOsd?: () => void;
  shouldShowOverlayLoadingOsd?: () => boolean;
  markOverlayLoadingOsdShown?: () => void;
  resetOverlayLoadingOsdSuppression?: () => void;
  resolveFallbackBounds?: () => WindowGeometry;
  hideNonNativeOverlayWhenTargetUnfocused?: boolean;
}): void {
  if (!args.mainWindow || args.mainWindow.isDestroyed()) {
    return;
  }

  const mainWindow = args.mainWindow;
  const overlayInteractionActive = args.overlayInteractionActive === true;

  if (args.modalActive) {
    if (args.isWindowsPlatform) {
      clearPendingWindowsOverlayReveal(mainWindow);
      setOverlayWindowOpacity(mainWindow, 0);
    }
    mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  if (args.suspendVisibleOverlay) {
    if (args.isWindowsPlatform) {
      clearPendingWindowsOverlayReveal(mainWindow);
      setOverlayWindowOpacity(mainWindow, 0);
    }
    mainWindow.setIgnoreMouseEvents(true, { forward: true });
    releaseOverlayWindowLevel(mainWindow);
    mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  const showPassiveVisibleOverlay = (): boolean => {
    const forceMousePassthrough = args.forceMousePassthrough === true;
    const wasVisible = mainWindow.isVisible();
    const isVisibleOverlayWindowFocused =
      typeof mainWindow.isFocused === 'function' && mainWindow.isFocused();
    const isVisibleOverlayFocused = overlayInteractionActive || isVisibleOverlayWindowFocused;
    const windowTracker = args.windowTracker;
    const canReportMacOSTargetMinimized =
      args.isMacOSPlatform && typeof windowTracker?.isTargetWindowMinimized === 'function';
    const isTrackedMacOSTargetMinimized =
      canReportMacOSTargetMinimized && windowTracker?.isTargetWindowMinimized() === true;
    const trackedMacOSTargetFocused = args.windowTracker?.isTargetWindowFocused?.();
    const shouldPreserveMacOSOverlayDuringForegroundProbe =
      args.isMacOSPlatform &&
      args.macOSForegroundProbeActive === true &&
      !!windowTracker &&
      !isTrackedMacOSTargetMinimized &&
      (windowTracker.isTracking() || windowTracker.getGeometry() !== null);
    const hasTransientMacOSTrackerLoss =
      args.isMacOSPlatform &&
      canReportMacOSTargetMinimized &&
      !!windowTracker &&
      !windowTracker.isTracking() &&
      !isTrackedMacOSTargetMinimized &&
      trackedMacOSTargetFocused !== false &&
      mainWindow.isVisible();
    const isTrackedMacOSTargetFocused =
      hasTransientMacOSTrackerLoss ||
      shouldPreserveMacOSOverlayDuringForegroundProbe ||
      !args.isMacOSPlatform ||
      !args.windowTracker
        ? true
        : (trackedMacOSTargetFocused ?? true);
    const shouldReleaseMacOSOverlayLevel =
      args.isMacOSPlatform &&
      !!args.windowTracker &&
      !hasTransientMacOSTrackerLoss &&
      !isVisibleOverlayFocused &&
      !isTrackedMacOSTargetFocused;
    // Renderer hover tracking temporarily disables this for subtitle and popup interaction.
    const shouldUseMacOSMousePassthrough = args.isMacOSPlatform && !overlayInteractionActive;
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
    const isNonNativeOverlay = !args.isWindowsPlatform && !args.isMacOSPlatform;
    const isNonNativePassiveOverlay = isNonNativeOverlay && !overlayInteractionActive;
    const hasNonNativeInputRegion =
      isNonNativePassiveOverlay && args.nonNativeInputRegionActive === true;
    const isTrackedNonNativeTargetFocused =
      !args.isWindowsPlatform && !args.isMacOSPlatform && !!args.windowTracker
        ? (args.windowTracker.isTargetWindowFocused?.() ?? true)
        : true;
    const shouldReleaseNonNativeOverlayLevel =
      isNonNativeOverlay &&
      !!args.windowTracker &&
      !isVisibleOverlayFocused &&
      !isTrackedNonNativeTargetFocused;
    const shouldIgnoreMouseEvents =
      shouldUseMacOSMousePassthrough ||
      forceMousePassthrough ||
      (isNonNativePassiveOverlay && !hasNonNativeInputRegion) ||
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

    if (shouldReleaseMacOSOverlayLevel) {
      releaseOverlayWindowLevel(mainWindow);
      if (wasVisible) {
        mainWindow.hide();
      }
      return false;
    }

    if (shouldBindTrackedWindowsOverlay) {
      // On Windows, z-order is enforced by the OS via the owner window mechanism
      // (SetWindowLongPtr GWLP_HWNDPARENT). The overlay is always above mpv
      // without any manual z-order management.
    } else if (shouldReleaseNonNativeOverlayLevel) {
      releaseOverlayWindowLevel(mainWindow);
      if (args.hideNonNativeOverlayWhenTargetUnfocused && wasVisible) {
        mainWindow.hide();
      }
    } else if (!forceMousePassthrough || args.isMacOSPlatform) {
      args.ensureOverlayWindowLevel(mainWindow);
    } else {
      releaseOverlayWindowLevel(mainWindow);
    }
    if (!wasVisible) {
      const hasWebContents =
        typeof (mainWindow as unknown as { webContents?: unknown }).webContents === 'object';
      if (
        hasWebContents &&
        !isOverlayWindowContentReady(mainWindow as unknown as import('electron').BrowserWindow)
      ) {
        // skip — ready-to-show hasn't fired yet; the onWindowContentReady
        // callback will trigger another visibility update when the renderer
        // has painted its first frame.
      } else if (
        ((args.isWindowsPlatform || args.isMacOSPlatform) && shouldIgnoreMouseEvents) ||
        isNonNativePassiveOverlay
      ) {
        if (args.isWindowsPlatform) {
          setOverlayWindowOpacity(mainWindow, 0);
        }
        mainWindow.showInactive();
        if (hasNonNativeInputRegion) {
          mainWindow.setIgnoreMouseEvents(false);
        } else {
          mainWindow.setIgnoreMouseEvents(true, { forward: true });
        }
        if (args.isWindowsPlatform) {
          scheduleWindowsOverlayReveal(
            mainWindow,
            shouldBindTrackedWindowsOverlay
              ? (window) => args.syncWindowsOverlayToMpvZOrder?.(window)
              : undefined,
          );
        }
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

    if (
      args.isMacOSPlatform &&
      overlayInteractionActive &&
      !forceMousePassthrough &&
      typeof mainWindow.isFocused === 'function' &&
      !mainWindow.isFocused()
    ) {
      mainWindow.focus();
    }

    return !shouldReleaseNonNativeOverlayLevel;
  };

  const shouldEnforceVisibleOverlayLayerOrder = (shouldEnforceLayerOrder: boolean): boolean =>
    shouldEnforceLayerOrder &&
    !args.isWindowsPlatform &&
    (!args.forceMousePassthrough || args.isMacOSPlatform === true);

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
  const maybeDismissOverlayLoadingOsd = (): void => {
    if (!args.isMacOSPlatform) {
      return;
    }
    args.dismissOverlayLoadingOsd?.();
  };

  const refreshNonNativeOverlayBoundsAfterFirstShow = (geometry: WindowGeometry | null): void => {
    if (
      geometry === null ||
      args.isMacOSPlatform ||
      args.isWindowsPlatform ||
      mainWindow.isVisible()
    ) {
      return;
    }
    if (pendingFirstShowBoundsRefreshGeometry.has(mainWindow)) {
      pendingFirstShowBoundsRefreshGeometry.set(mainWindow, geometry);
      return;
    }
    pendingFirstShowBoundsRefreshGeometry.set(mainWindow, geometry);
    mainWindow.once('show', () => {
      const pendingGeometry = pendingFirstShowBoundsRefreshGeometry.get(mainWindow);
      pendingFirstShowBoundsRefreshGeometry.delete(mainWindow);
      if (mainWindow.isDestroyed() || !mainWindow.isVisible()) {
        return;
      }
      if (pendingGeometry) {
        args.updateVisibleOverlayBounds(pendingGeometry);
      }
    });
  };

  if (!args.visibleOverlayVisible) {
    args.setTrackerNotReadyWarningShown(false);
    args.resetOverlayLoadingOsdSuppression?.();
    maybeDismissOverlayLoadingOsd();
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
    maybeDismissOverlayLoadingOsd();
    const geometry = args.windowTracker.getGeometry();
    if (geometry) {
      args.updateVisibleOverlayBounds(geometry);
      refreshNonNativeOverlayBoundsAfterFirstShow(geometry);
    }
    args.syncPrimaryOverlayWindowLayer('visible');
    const shouldEnforceLayerOrder = showPassiveVisibleOverlay();
    if (shouldEnforceVisibleOverlayLayerOrder(shouldEnforceLayerOrder)) {
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
    mainWindow.setIgnoreMouseEvents(true, { forward: true });
    releaseOverlayWindowLevel(mainWindow);
    mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  const hasRetainedTrackedGeometry = args.windowTracker.getGeometry() !== null;
  const hasActiveMacOSTargetSignal =
    args.isMacOSPlatform && (args.windowTracker.isTargetWindowFocused?.() ?? false);
  const hasActiveMacOSOverlaySignal = args.isMacOSPlatform && overlayInteractionActive;
  const canReportMacOSTargetMinimized =
    args.isMacOSPlatform && typeof args.windowTracker.isTargetWindowMinimized === 'function';
  const isTrackedMacOSTargetMinimized =
    canReportMacOSTargetMinimized && args.windowTracker.isTargetWindowMinimized();
  const shouldPreserveTransientTrackedOverlay =
    (args.isMacOSPlatform &&
      !isTrackedMacOSTargetMinimized &&
      (hasRetainedTrackedGeometry ||
        (mainWindow.isVisible() && hasActiveMacOSOverlaySignal) ||
        (mainWindow.isVisible() && hasActiveMacOSTargetSignal) ||
        (canReportMacOSTargetMinimized && mainWindow.isVisible()))) ||
    (args.isWindowsPlatform &&
      typeof args.windowTracker.isTargetWindowMinimized === 'function' &&
      !args.windowTracker.isTargetWindowMinimized());

  if (
    shouldPreserveTransientTrackedOverlay &&
    (mainWindow.isVisible() || hasRetainedTrackedGeometry)
  ) {
    args.setTrackerNotReadyWarningShown(false);
    maybeDismissOverlayLoadingOsd();
    const geometry = args.windowTracker.getGeometry();
    if (geometry) {
      args.updateVisibleOverlayBounds(geometry);
    }
    args.syncPrimaryOverlayWindowLayer('visible');
    const shouldEnforceLayerOrder = showPassiveVisibleOverlay();
    if (shouldEnforceVisibleOverlayLayerOrder(shouldEnforceLayerOrder)) {
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
