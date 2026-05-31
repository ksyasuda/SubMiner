import type { OverlayContentMeasurement, OverlayContentRect } from '../../types';
import type { AutoplayReadySignal } from './autoplay-ready-gate';

export type VisibleOverlayAutoplayReadinessDeps = {
  getVisibleOverlayVisible: () => boolean;
  isOverlayWindowReady: () => boolean;
  getLatestVisibleMeasurement: () => OverlayContentMeasurement | null;
};

function hasArea(rect: OverlayContentRect): boolean {
  return rect.width > 0 && rect.height > 0;
}

function hasMeasuredInteractiveContent(measurement: OverlayContentMeasurement): boolean {
  const rects =
    Array.isArray(measurement.interactiveRects) && measurement.interactiveRects.some(hasArea)
      ? measurement.interactiveRects
      : measurement.contentRect
        ? [measurement.contentRect]
        : [];

  return rects.some(hasArea);
}

export function isVisibleOverlayAutoplayTargetReady(
  deps: VisibleOverlayAutoplayReadinessDeps,
  signal: AutoplayReadySignal,
): boolean {
  if (!deps.getVisibleOverlayVisible()) {
    return true;
  }

  const subtitleText = signal.payload.text.trim();
  if (!subtitleText || subtitleText === '__warm__') {
    return false;
  }

  if (!deps.isOverlayWindowReady()) {
    return false;
  }

  const measurement = deps.getLatestVisibleMeasurement();
  if (!measurement || measurement.measuredAtMs < signal.requestedAtMs) {
    return false;
  }

  return hasMeasuredInteractiveContent(measurement);
}
