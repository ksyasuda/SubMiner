import type { WindowGeometry } from '../../types';

export function createUpdateVisibleOverlayBoundsHandler(deps: {
  setOverlayWindowBounds: (geometry: WindowGeometry) => void;
  afterSetOverlayWindowBounds?: (geometry: WindowGeometry) => void;
}) {
  return (geometry: WindowGeometry): void => {
    deps.setOverlayWindowBounds(geometry);
    deps.afterSetOverlayWindowBounds?.(geometry);
  };
}

export function createEnsureOverlayWindowLevelHandler(deps: {
  shouldSuppressOverlayWindowLevel?: (window: unknown) => boolean;
  ensureOverlayWindowLevelCore: (window: unknown) => void;
  afterEnsureOverlayWindowLevel?: (window: unknown) => void;
}) {
  return (window: unknown): void => {
    if (deps.shouldSuppressOverlayWindowLevel?.(window) === true) {
      return;
    }
    deps.ensureOverlayWindowLevelCore(window);
    deps.afterEnsureOverlayWindowLevel?.(window);
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
