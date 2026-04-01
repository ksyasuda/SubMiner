import type { WindowGeometry } from '../types';
import type { OverlayGeometryRuntime } from './overlay-geometry-runtime';

export function createOverlayGeometryAccessors(deps: {
  getOverlayGeometryRuntime: () => OverlayGeometryRuntime<any> | null;
  getWindowTracker: () => { getGeometry?: () => WindowGeometry | null } | null;
  screen: {
    getCursorScreenPoint: () => { x: number; y: number };
    getDisplayNearestPoint: (point: { x: number; y: number }) => {
      workArea: { x: number; y: number; width: number; height: number };
    };
  };
}) {
  const getOverlayGeometryFallback = (): WindowGeometry => {
    const runtime = deps.getOverlayGeometryRuntime();
    if (runtime) {
      return runtime.getOverlayGeometryFallback();
    }

    const cursorPoint = deps.screen.getCursorScreenPoint();
    const display = deps.screen.getDisplayNearestPoint(cursorPoint);
    const bounds = display.workArea;
    return {
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
    };
  };

  const getCurrentOverlayGeometry = (): WindowGeometry => {
    const runtime = deps.getOverlayGeometryRuntime();
    if (runtime) {
      return runtime.getCurrentOverlayGeometry();
    }

    const trackerGeometry = deps.getWindowTracker()?.getGeometry?.() ?? null;
    if (trackerGeometry) {
      return trackerGeometry;
    }

    return getOverlayGeometryFallback();
  };

  const geometryMatches = (a: WindowGeometry | null, b: WindowGeometry | null): boolean => {
    const runtime = deps.getOverlayGeometryRuntime();
    if (runtime) {
      return runtime.geometryMatches(a, b);
    }

    if (!a || !b) {
      return false;
    }

    return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
  };

  return {
    getOverlayGeometryFallback,
    getCurrentOverlayGeometry,
    geometryMatches,
  };
}
