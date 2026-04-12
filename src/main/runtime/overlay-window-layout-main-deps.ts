import type {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateVisibleOverlayBoundsHandler,
} from './overlay-window-layout';

type UpdateVisibleOverlayBoundsMainDeps = Parameters<
  typeof createUpdateVisibleOverlayBoundsHandler
>[0];
type EnsureOverlayWindowLevelMainDeps = Parameters<typeof createEnsureOverlayWindowLevelHandler>[0];
type EnforceOverlayLayerOrderMainDeps = Parameters<typeof createEnforceOverlayLayerOrderHandler>[0];

export function createBuildUpdateVisibleOverlayBoundsMainDepsHandler(
  deps: UpdateVisibleOverlayBoundsMainDeps,
) {
  return (): UpdateVisibleOverlayBoundsMainDeps => ({
    setOverlayWindowBounds: (geometry) => deps.setOverlayWindowBounds(geometry),
    afterSetOverlayWindowBounds: (geometry) => deps.afterSetOverlayWindowBounds?.(geometry),
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
    getMainWindow: () => deps.getMainWindow(),
    ensureOverlayWindowLevel: (window: unknown) => deps.ensureOverlayWindowLevel(window),
  });
}
