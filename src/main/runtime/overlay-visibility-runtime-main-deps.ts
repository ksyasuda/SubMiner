import type { BrowserWindow } from 'electron';
import type { WindowGeometry } from '../../types';
import type { OverlayVisibilityRuntimeDeps } from '../overlay-visibility-runtime';

export function createBuildOverlayVisibilityRuntimeMainDepsHandler(
  deps: OverlayVisibilityRuntimeDeps,
) {
  return (): OverlayVisibilityRuntimeDeps => ({
    getMainWindow: () => deps.getMainWindow(),
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    getWindowTracker: () => deps.getWindowTracker(),
    getTrackerNotReadyWarningShown: () => deps.getTrackerNotReadyWarningShown(),
    setTrackerNotReadyWarningShown: (shown: boolean) => deps.setTrackerNotReadyWarningShown(shown),
    updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
      deps.updateVisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window: BrowserWindow) => deps.ensureOverlayWindowLevel(window),
    syncPrimaryOverlayWindowLayer: (layer: 'visible') => deps.syncPrimaryOverlayWindowLayer(layer),
    enforceOverlayLayerOrder: () => deps.enforceOverlayLayerOrder(),
    syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
    isMacOSPlatform: () => deps.isMacOSPlatform(),
    showOverlayLoadingOsd: (message: string) => deps.showOverlayLoadingOsd(message),
    resolveFallbackBounds: () => deps.resolveFallbackBounds(),
  });
}
