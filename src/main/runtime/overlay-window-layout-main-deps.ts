import type {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateInvisibleOverlayBoundsHandler,
  createUpdateVisibleOverlayBoundsHandler,
} from './overlay-window-layout';

type UpdateVisibleOverlayBoundsMainDeps = Parameters<typeof createUpdateVisibleOverlayBoundsHandler>[0];
type UpdateInvisibleOverlayBoundsMainDeps = Parameters<typeof createUpdateInvisibleOverlayBoundsHandler>[0];
type EnsureOverlayWindowLevelMainDeps = Parameters<typeof createEnsureOverlayWindowLevelHandler>[0];
type EnforceOverlayLayerOrderMainDeps = Parameters<typeof createEnforceOverlayLayerOrderHandler>[0];

export function createBuildUpdateVisibleOverlayBoundsMainDepsHandler(
  deps: UpdateVisibleOverlayBoundsMainDeps,
) {
  return (): UpdateVisibleOverlayBoundsMainDeps => ({
    setOverlayWindowBounds: (layer, geometry) => deps.setOverlayWindowBounds(layer, geometry),
  });
}

export function createBuildUpdateInvisibleOverlayBoundsMainDepsHandler(
  deps: UpdateInvisibleOverlayBoundsMainDeps,
) {
  return (): UpdateInvisibleOverlayBoundsMainDeps => ({
    setOverlayWindowBounds: (layer, geometry) => deps.setOverlayWindowBounds(layer, geometry),
  });
}

export function createBuildEnsureOverlayWindowLevelMainDepsHandler(
  deps: EnsureOverlayWindowLevelMainDeps,
) {
  return (): EnsureOverlayWindowLevelMainDeps => ({
    ensureOverlayWindowLevelCore: (window: unknown) => deps.ensureOverlayWindowLevelCore(window),
  });
}

export function createBuildEnforceOverlayLayerOrderMainDepsHandler(
  deps: EnforceOverlayLayerOrderMainDeps,
) {
  return (): EnforceOverlayLayerOrderMainDeps => ({
    enforceOverlayLayerOrderCore: (params) => deps.enforceOverlayLayerOrderCore(params),
    getVisibleOverlayVisible: () => deps.getVisibleOverlayVisible(),
    getInvisibleOverlayVisible: () => deps.getInvisibleOverlayVisible(),
    getMainWindow: () => deps.getMainWindow(),
    getInvisibleWindow: () => deps.getInvisibleWindow(),
    ensureOverlayWindowLevel: (window: unknown) => deps.ensureOverlayWindowLevel(window),
  });
}
