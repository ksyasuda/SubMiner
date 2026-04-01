import {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateVisibleOverlayBoundsHandler,
} from './runtime/overlay-window-layout';
import {
  createBuildEnforceOverlayLayerOrderMainDepsHandler,
  createBuildEnsureOverlayWindowLevelMainDepsHandler,
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler,
} from './runtime/overlay-window-layout-main-deps';
import type { WindowGeometry } from '../types';

type BrowserWindowLike = {
  isDestroyed: () => boolean;
};

type ScreenLike = {
  getCursorScreenPoint: () => { x: number; y: number };
  getDisplayNearestPoint: (point: { x: number; y: number }) => {
    workArea: { x: number; y: number; width: number; height: number };
  };
};

export interface OverlayGeometryWindowState<TWindow extends BrowserWindowLike = BrowserWindowLike> {
  getMainWindow: () => TWindow | null;
  setOverlayWindowBounds: (geometry: WindowGeometry) => void;
  setModalWindowBounds: (geometry: WindowGeometry) => void;
  getVisibleOverlayVisible: () => boolean;
}

export interface OverlayGeometryInput<TWindow extends BrowserWindowLike = BrowserWindowLike> {
  screen: ScreenLike;
  windowState: OverlayGeometryWindowState<TWindow>;
  getWindowTracker: () => { getGeometry?: () => WindowGeometry | null } | null;
  ensureOverlayWindowLevelCore: (window: TWindow) => void;
  syncOverlayWindowLayer: (window: TWindow, layer: 'visible') => void;
  enforceOverlayLayerOrderCore: (params: {
    visibleOverlayVisible: boolean;
    mainWindow: TWindow | null;
    ensureOverlayWindowLevel: (window: TWindow) => void;
  }) => void;
}

export interface OverlayGeometryRuntime<TWindow extends BrowserWindowLike = BrowserWindowLike> {
  getLastOverlayWindowGeometry: () => WindowGeometry | null;
  getOverlayGeometryFallback: () => WindowGeometry;
  getCurrentOverlayGeometry: () => WindowGeometry;
  geometryMatches: (a: WindowGeometry | null, b: WindowGeometry | null) => boolean;
  applyOverlayRegions: (geometry: WindowGeometry) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  syncPrimaryOverlayWindowLayer: (layer: 'visible') => void;
  enforceOverlayLayerOrder: () => void;
}

export function createOverlayGeometryRuntime<TWindow extends BrowserWindowLike = BrowserWindowLike>(
  input: OverlayGeometryInput<TWindow>,
): OverlayGeometryRuntime<TWindow> {
  let lastOverlayWindowGeometry: WindowGeometry | null = null;

  const getOverlayGeometryFallback = (): WindowGeometry => {
    const cursorPoint = input.screen.getCursorScreenPoint();
    const display = input.screen.getDisplayNearestPoint(cursorPoint);
    const bounds = display.workArea;
    return {
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
    };
  };

  const getCurrentOverlayGeometry = (): WindowGeometry => {
    if (lastOverlayWindowGeometry) return lastOverlayWindowGeometry;
    const trackerGeometry = input.getWindowTracker()?.getGeometry?.() ?? null;
    if (trackerGeometry) return trackerGeometry;
    return getOverlayGeometryFallback();
  };

  const geometryMatches = (a: WindowGeometry | null, b: WindowGeometry | null): boolean => {
    if (!a || !b) return false;
    return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
  };

  const applyOverlayRegions = (geometry: WindowGeometry): void => {
    lastOverlayWindowGeometry = geometry;
    input.windowState.setOverlayWindowBounds(geometry);
    input.windowState.setModalWindowBounds(geometry);
  };

  const updateVisibleOverlayBounds = createUpdateVisibleOverlayBoundsHandler(
    createBuildUpdateVisibleOverlayBoundsMainDepsHandler({
      setOverlayWindowBounds: (geometry) => applyOverlayRegions(geometry),
    })(),
  );

  const ensureOverlayWindowLevel = createEnsureOverlayWindowLevelHandler(
    createBuildEnsureOverlayWindowLevelMainDepsHandler({
      ensureOverlayWindowLevelCore: (window) =>
        input.ensureOverlayWindowLevelCore(window as TWindow),
    })(),
  );

  const syncPrimaryOverlayWindowLayer = (layer: 'visible'): void => {
    const mainWindow = input.windowState.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) return;
    input.syncOverlayWindowLayer(mainWindow, layer);
  };

  const enforceOverlayLayerOrder = createEnforceOverlayLayerOrderHandler(
    createBuildEnforceOverlayLayerOrderMainDepsHandler({
      enforceOverlayLayerOrderCore: (params) =>
        input.enforceOverlayLayerOrderCore({
          visibleOverlayVisible: params.visibleOverlayVisible,
          mainWindow: params.mainWindow as TWindow | null,
          ensureOverlayWindowLevel: (window) => params.ensureOverlayWindowLevel(window as TWindow),
        }),
      getVisibleOverlayVisible: () => input.windowState.getVisibleOverlayVisible(),
      getMainWindow: () => input.windowState.getMainWindow(),
      ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window as TWindow),
    })(),
  );

  return {
    getLastOverlayWindowGeometry: () => lastOverlayWindowGeometry,
    getOverlayGeometryFallback,
    getCurrentOverlayGeometry,
    geometryMatches,
    applyOverlayRegions,
    updateVisibleOverlayBounds,
    ensureOverlayWindowLevel,
    syncPrimaryOverlayWindowLayer,
    enforceOverlayLayerOrder,
  };
}
