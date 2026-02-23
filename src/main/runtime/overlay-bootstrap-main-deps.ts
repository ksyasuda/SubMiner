import type { OverlayWindowResolver } from '../overlay-runtime';

type OverlayContentMeasurementStoreMainDeps = {
  now: () => number;
  warn: (message: string) => void;
};

export function createBuildOverlayContentMeasurementStoreMainDepsHandler(
  deps: OverlayContentMeasurementStoreMainDeps,
) {
  return (): OverlayContentMeasurementStoreMainDeps => ({
    now: () => deps.now(),
    warn: (message: string) => deps.warn(message),
  });
}

export function createBuildOverlayModalRuntimeMainDepsHandler(deps: OverlayWindowResolver) {
  return (): OverlayWindowResolver => ({
    getMainWindow: () => deps.getMainWindow(),
    getInvisibleWindow: () => deps.getInvisibleWindow(),
  });
}
