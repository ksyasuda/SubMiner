import type { OverlayContentMeasurement, OverlayContentRect } from '../../types';

type PointerPoint = { x: number; y: number };
type PointerRect = { x: number; y: number; width: number; height: number };

type PointerInteractionWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  getBounds: () => PointerRect;
};

export type WindowsOverlayPointerInteractionDeps = {
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => PointerInteractionWindow | null;
  getCursorScreenPoint: () => PointerPoint;
  getSubtitleMeasurement: () => OverlayContentMeasurement | null;
  getRendererInteractiveHint?: () => boolean;
  /** True when a modal/stats/separate window owns input. */
  shouldSuspend: () => boolean;
  getInteractionActive: () => boolean;
  setInteractionActive: (active: boolean) => void;
};

// Match Linux fallback padding so hover survives tiny measurement/cursor gaps.
const SUBTITLE_HIT_PADDING_PX = 6;

function measuredRectsForInput(
  measurement: OverlayContentMeasurement | null,
): OverlayContentRect[] {
  if (!measurement) return [];
  return Array.isArray(measurement.interactiveRects) && measurement.interactiveRects.length > 0
    ? measurement.interactiveRects
    : measurement.contentRect
      ? [measurement.contentRect]
      : [];
}

function isCursorOverRect(
  cursor: PointerPoint,
  bounds: PointerRect,
  viewport: { width: number; height: number },
  rect: OverlayContentRect,
): boolean {
  if (!(bounds.width > 0) || !(bounds.height > 0)) return false;
  if (!(viewport.width > 0) || !(viewport.height > 0)) return false;
  if (!(rect.width > 0) || !(rect.height > 0)) return false;

  const scaleX = bounds.width / viewport.width;
  const scaleY = bounds.height / viewport.height;
  const left = bounds.x + rect.x * scaleX - SUBTITLE_HIT_PADDING_PX;
  const top = bounds.y + rect.y * scaleY - SUBTITLE_HIT_PADDING_PX;
  const right = left + rect.width * scaleX + SUBTITLE_HIT_PADDING_PX * 2;
  const bottom = top + rect.height * scaleY + SUBTITLE_HIT_PADDING_PX * 2;

  return cursor.x >= left && cursor.x <= right && cursor.y >= top && cursor.y <= bottom;
}

export function isCursorOverWindowsOverlayInteractiveRect(
  cursor: PointerPoint,
  bounds: PointerRect,
  measurement: OverlayContentMeasurement | null,
): boolean {
  if (!measurement) return false;
  return measuredRectsForInput(measurement).some((rect) =>
    isCursorOverRect(cursor, bounds, measurement.viewport, rect),
  );
}

/**
 * Returns the desired Windows overlay mouse-input state, or null when another surface
 * currently owns interaction and the fallback should not touch BrowserWindow passthrough.
 */
export function resolveDesiredWindowsOverlayInteractive(
  deps: WindowsOverlayPointerInteractionDeps,
): boolean | null {
  if (!deps.getVisibleOverlayVisible()) return false;
  if (deps.shouldSuspend()) return null;

  const mainWindow = deps.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    return null;
  }

  if (deps.getRendererInteractiveHint?.()) return true;
  return isCursorOverWindowsOverlayInteractiveRect(
    deps.getCursorScreenPoint(),
    mainWindow.getBounds(),
    deps.getSubtitleMeasurement(),
  );
}

export function tickWindowsOverlayPointerInteraction(
  deps: WindowsOverlayPointerInteractionDeps,
): void {
  const desired = resolveDesiredWindowsOverlayInteractive(deps);
  if (desired === null) return;
  if (deps.getInteractionActive() === desired) return;
  deps.setInteractionActive(desired);
}
