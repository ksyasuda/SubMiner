import type { BrowserWindow } from 'electron';
import type { WindowGeometry } from '../../types';
import type { OverlayVisibilityRuntimeDeps } from '../overlay-visibility-runtime';

export function createBuildOverlayVisibilityRuntimeMainDepsHandler(deps: {
  getMainWindow: () => BrowserWindow | null;
  getInvisibleWindow: () => BrowserWindow | null;
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  getWindowTracker: () => unknown | null;
  getTrackerNotReadyWarningShown: () => boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  updateInvisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
}) {
  return (): OverlayVisibilityRuntimeDeps => ({
    getMainWindow: () => deps.getMainWindow(),
    getInvisibleWindow: () => deps.getInvisibleWindow(),
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    getInvisibleOverlayVisible: () => deps.getInvisibleOverlayVisible(),
    getWindowTracker: () => deps.getWindowTracker() as never,
    getTrackerNotReadyWarningShown: () => deps.getTrackerNotReadyWarningShown(),
    setTrackerNotReadyWarningShown: (shown: boolean) => deps.setTrackerNotReadyWarningShown(shown),
    updateVisibleOverlayBounds: (geometry: WindowGeometry) => deps.updateVisibleOverlayBounds(geometry),
    updateInvisibleOverlayBounds: (geometry: WindowGeometry) =>
      deps.updateInvisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window: BrowserWindow) => deps.ensureOverlayWindowLevel(window),
    enforceOverlayLayerOrder: () => deps.enforceOverlayLayerOrder(),
    syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
  });
}
