import type { BrowserWindow } from 'electron';

import type { BaseWindowTracker } from '../window-trackers';
import type { WindowGeometry } from '../types';
import { updateVisibleOverlayVisibility } from '../core/services';

const OVERLAY_LOADING_OSD_COOLDOWN_MS = 30_000;

export interface OverlayVisibilityRuntimeDeps {
  getMainWindow: () => BrowserWindow | null;
  getModalActive: () => boolean;
  getVisibleOverlayVisible: () => boolean;
  getForceMousePassthrough: () => boolean;
  getNonNativeInputRegionActive?: () => boolean;
  getSuspendVisibleOverlay?: () => boolean;
  getOverlayInteractionActive?: () => boolean;
  getWindowTracker: () => BaseWindowTracker | null;
  getLastKnownWindowsForegroundProcessName?: () => string | null;
  getWindowsOverlayProcessName?: () => string | null;
  getWindowsFocusHandoffGraceActive?: () => boolean;
  getMacOSForegroundProbeActive?: () => boolean;
  getTrackerNotReadyWarningShown: () => boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  syncWindowsOverlayToMpvZOrder?: (window: BrowserWindow) => void;
  syncPrimaryOverlayWindowLayer: (layer: 'visible') => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
  isMacOSPlatform: () => boolean;
  isWindowsPlatform: () => boolean;
  showOverlayLoadingOsd: (message: string) => void;
  dismissOverlayLoadingOsd?: () => void;
  resolveFallbackBounds: () => WindowGeometry;
  hideNonNativeOverlayWhenTargetUnfocused?: () => boolean;
}

export interface OverlayVisibilityRuntimeService {
  updateVisibleOverlayVisibility: () => void;
}

export function createOverlayVisibilityRuntimeService(
  deps: OverlayVisibilityRuntimeDeps,
): OverlayVisibilityRuntimeService {
  let lastOverlayLoadingOsdAtMs: number | null = null;

  return {
    updateVisibleOverlayVisibility(): void {
      const visibleOverlayVisible = deps.getVisibleOverlayVisible();
      const forceMousePassthrough = deps.getForceMousePassthrough();
      const suspendVisibleOverlay = deps.getSuspendVisibleOverlay?.() ?? false;
      const windowTracker = deps.getWindowTracker();
      const mainWindow = deps.getMainWindow();

      updateVisibleOverlayVisibility({
        visibleOverlayVisible,
        modalActive: deps.getModalActive(),
        forceMousePassthrough,
        nonNativeInputRegionActive: deps.getNonNativeInputRegionActive?.() ?? false,
        suspendVisibleOverlay,
        overlayInteractionActive: deps.getOverlayInteractionActive?.() ?? false,
        mainWindow,
        windowTracker,
        lastKnownWindowsForegroundProcessName: deps.getLastKnownWindowsForegroundProcessName?.(),
        windowsOverlayProcessName: deps.getWindowsOverlayProcessName?.() ?? null,
        windowsFocusHandoffGraceActive: deps.getWindowsFocusHandoffGraceActive?.() ?? false,
        macOSForegroundProbeActive: deps.getMacOSForegroundProbeActive?.() ?? false,
        trackerNotReadyWarningShown: deps.getTrackerNotReadyWarningShown(),
        setTrackerNotReadyWarningShown: (shown: boolean) => {
          deps.setTrackerNotReadyWarningShown(shown);
        },
        updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
          deps.updateVisibleOverlayBounds(geometry),
        ensureOverlayWindowLevel: (window: BrowserWindow) => deps.ensureOverlayWindowLevel(window),
        syncWindowsOverlayToMpvZOrder: (window: BrowserWindow) =>
          deps.syncWindowsOverlayToMpvZOrder?.(window),
        syncPrimaryOverlayWindowLayer: (layer: 'visible') =>
          deps.syncPrimaryOverlayWindowLayer(layer),
        enforceOverlayLayerOrder: () => deps.enforceOverlayLayerOrder(),
        syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
        isMacOSPlatform: deps.isMacOSPlatform(),
        isWindowsPlatform: deps.isWindowsPlatform(),
        showOverlayLoadingOsd: (message: string) => deps.showOverlayLoadingOsd(message),
        dismissOverlayLoadingOsd: () => deps.dismissOverlayLoadingOsd?.(),
        shouldShowOverlayLoadingOsd: () =>
          lastOverlayLoadingOsdAtMs === null ||
          Date.now() - lastOverlayLoadingOsdAtMs >= OVERLAY_LOADING_OSD_COOLDOWN_MS,
        markOverlayLoadingOsdShown: () => {
          lastOverlayLoadingOsdAtMs = Date.now();
        },
        resetOverlayLoadingOsdSuppression: () => {
          lastOverlayLoadingOsdAtMs = null;
        },
        hideNonNativeOverlayWhenTargetUnfocused:
          deps.hideNonNativeOverlayWhenTargetUnfocused?.() ?? false,
        resolveFallbackBounds: () => deps.resolveFallbackBounds(),
      });
    },
  };
}
