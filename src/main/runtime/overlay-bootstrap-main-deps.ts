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

export function createBuildOverlayModalRuntimeMainDepsHandler(
  deps: OverlayWindowResolver,
) {
  return (): OverlayWindowResolver => ({
    getMainWindow: () => deps.getMainWindow(),
    getModalWindow: () => deps.getModalWindow(),
    createModalWindow: () => deps.createModalWindow(),
    getModalGeometry: () => deps.getModalGeometry(),
    setModalWindowBounds: (geometry) => deps.setModalWindowBounds(geometry),
  });
}
