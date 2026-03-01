import type { WindowGeometry } from '../../types';

export function createUpdateVisibleOverlayBoundsHandler(deps: {
  setOverlayWindowBounds: (geometry: WindowGeometry) => void;
}) {
  return (geometry: WindowGeometry): void => {
    deps.setOverlayWindowBounds(geometry);
  };
}

export function createEnsureOverlayWindowLevelHandler(deps: {
  ensureOverlayWindowLevelCore: (window: unknown) => void;
}) {
  return (window: unknown): void => {
    deps.ensureOverlayWindowLevelCore(window);
  };
}

export function createEnforceOverlayLayerOrderHandler(deps: {
  enforceOverlayLayerOrderCore: (params: {
    visibleOverlayVisible: boolean;
    mainWindow: unknown;
    ensureOverlayWindowLevel: (window: unknown) => void;
  }) => void;
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => unknown;
  ensureOverlayWindowLevel: (window: unknown) => void;
}) {
  return (): void => {
    deps.enforceOverlayLayerOrderCore({
      visibleOverlayVisible: deps.getVisibleOverlayVisible(),
      mainWindow: deps.getMainWindow(),
      ensureOverlayWindowLevel: deps.ensureOverlayWindowLevel,
    });
  };
}
