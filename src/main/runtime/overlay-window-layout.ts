import type { WindowGeometry } from '../../types';

export function createUpdateVisibleOverlayBoundsHandler(deps: {
  setOverlayWindowBounds: (layer: 'visible' | 'invisible', geometry: WindowGeometry) => void;
}) {
  return (geometry: WindowGeometry): void => {
    deps.setOverlayWindowBounds('visible', geometry);
  };
}

export function createUpdateInvisibleOverlayBoundsHandler(deps: {
  setOverlayWindowBounds: (layer: 'visible' | 'invisible', geometry: WindowGeometry) => void;
}) {
  return (geometry: WindowGeometry): void => {
    deps.setOverlayWindowBounds('invisible', geometry);
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
    invisibleOverlayVisible: boolean;
    mainWindow: unknown;
    invisibleWindow: unknown;
    ensureOverlayWindowLevel: (window: unknown) => void;
  }) => void;
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  getMainWindow: () => unknown;
  getInvisibleWindow: () => unknown;
  ensureOverlayWindowLevel: (window: unknown) => void;
}) {
  return (): void => {
    deps.enforceOverlayLayerOrderCore({
      visibleOverlayVisible: deps.getVisibleOverlayVisible(),
      invisibleOverlayVisible: deps.getInvisibleOverlayVisible(),
      mainWindow: deps.getMainWindow(),
      invisibleWindow: deps.getInvisibleWindow(),
      ensureOverlayWindowLevel: deps.ensureOverlayWindowLevel,
    });
  };
}
