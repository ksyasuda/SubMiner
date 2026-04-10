import type { BrowserWindow } from 'electron';
import type { WindowGeometry } from '../../types';
import type { OverlayVisibilityRuntimeDeps } from '../overlay-visibility-runtime';

export function createBuildOverlayVisibilityRuntimeMainDepsHandler(
  deps: OverlayVisibilityRuntimeDeps,
) {
  return (): OverlayVisibilityRuntimeDeps => ({
    getMainWindow: () => deps.getMainWindow(),
    getModalActive: () => deps.getModalActive(),
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    getForceMousePassthrough: () => deps.getForceMousePassthrough(),
    getWindowTracker: () => deps.getWindowTracker(),
    getLastKnownWindowsForegroundProcessName: () =>
      deps.getLastKnownWindowsForegroundProcessName?.() ?? null,
    getWindowsOverlayProcessName: () => deps.getWindowsOverlayProcessName?.() ?? null,
    getWindowsFocusHandoffGraceActive: () => deps.getWindowsFocusHandoffGraceActive?.() ?? false,
    getTrackerNotReadyWarningShown: () => deps.getTrackerNotReadyWarningShown(),
    setTrackerNotReadyWarningShown: (shown: boolean) => deps.setTrackerNotReadyWarningShown(shown),
    updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
      deps.updateVisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window: BrowserWindow) => deps.ensureOverlayWindowLevel(window),
    syncWindowsOverlayToMpvZOrder: (window: BrowserWindow) =>
      deps.syncWindowsOverlayToMpvZOrder?.(window),
    syncPrimaryOverlayWindowLayer: (layer: 'visible') => deps.syncPrimaryOverlayWindowLayer(layer),
    enforceOverlayLayerOrder: () => deps.enforceOverlayLayerOrder(),
    syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
    isMacOSPlatform: () => deps.isMacOSPlatform(),
    isWindowsPlatform: () => deps.isWindowsPlatform(),
    showOverlayLoadingOsd: (message: string) => deps.showOverlayLoadingOsd(message),
    resolveFallbackBounds: () => deps.resolveFallbackBounds(),
  });
}
