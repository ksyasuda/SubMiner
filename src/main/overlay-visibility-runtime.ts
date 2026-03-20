import type { BrowserWindow } from 'electron';

import type { BaseWindowTracker } from '../window-trackers';
import type { WindowGeometry } from '../types';
import { updateVisibleOverlayVisibility } from '../core/services';

const OVERLAY_LOADING_OSD_COOLDOWN_MS = 30_000;

export interface OverlayVisibilityRuntimeDeps {
  getMainWindow: () => BrowserWindow | null;
  getVisibleOverlayVisible: () => boolean;
  getForceMousePassthrough: () => boolean;
  getWindowTracker: () => BaseWindowTracker | null;
  getTrackerNotReadyWarningShown: () => boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  syncPrimaryOverlayWindowLayer: (layer: 'visible') => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
  isMacOSPlatform: () => boolean;
  isWindowsPlatform: () => boolean;
  showOverlayLoadingOsd: (message: string) => void;
  resolveFallbackBounds: () => WindowGeometry;
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
      updateVisibleOverlayVisibility({
        visibleOverlayVisible: deps.getVisibleOverlayVisible(),
        forceMousePassthrough: deps.getForceMousePassthrough(),
        mainWindow: deps.getMainWindow(),
        windowTracker: deps.getWindowTracker(),
        trackerNotReadyWarningShown: deps.getTrackerNotReadyWarningShown(),
        setTrackerNotReadyWarningShown: (shown: boolean) => {
          deps.setTrackerNotReadyWarningShown(shown);
        },
        updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
          deps.updateVisibleOverlayBounds(geometry),
        ensureOverlayWindowLevel: (window: BrowserWindow) => deps.ensureOverlayWindowLevel(window),
        syncPrimaryOverlayWindowLayer: (layer: 'visible') =>
          deps.syncPrimaryOverlayWindowLayer(layer),
        enforceOverlayLayerOrder: () => deps.enforceOverlayLayerOrder(),
        syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
        isMacOSPlatform: deps.isMacOSPlatform(),
        isWindowsPlatform: deps.isWindowsPlatform(),
        showOverlayLoadingOsd: (message: string) => deps.showOverlayLoadingOsd(message),
        shouldShowOverlayLoadingOsd: () =>
          lastOverlayLoadingOsdAtMs === null ||
          Date.now() - lastOverlayLoadingOsdAtMs >= OVERLAY_LOADING_OSD_COOLDOWN_MS,
        markOverlayLoadingOsdShown: () => {
          lastOverlayLoadingOsdAtMs = Date.now();
        },
        resetOverlayLoadingOsdSuppression: () => {
          lastOverlayLoadingOsdAtMs = null;
        },
        resolveFallbackBounds: () => deps.resolveFallbackBounds(),
      });
    },
  };
}
