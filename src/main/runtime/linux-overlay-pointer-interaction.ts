/*
  Linux overlay pointer-interaction loop.

  Electron cannot forward mouse-move events through a click-through window on Linux/X11
  (the `forward` option of setIgnoreMouseEvents is unsupported there — electron/electron#16777).
  The overlay's hover/lookup interaction relied on those forwarded events, so under XWayland
  the click-through overlay never sees the cursor and stays inert.

  This restores the Windows/macOS behavior with a main-process cursor poll: while the overlay
  is shown, it tracks the global cursor against the reported subtitle rect and flips the overlay
  interactive when the cursor is over a subtitle (or when the renderer reports an interactive
  region such as an open Yomitan popup or modal), and back to click-through otherwise.
*/

import type { OverlayContentMeasurement } from '../../types';

export type PointerPoint = { x: number; y: number };
export type PointerRect = { x: number; y: number; width: number; height: number };
export type PointerViewport = { width: number; height: number };

export type OverlayContentMeasurementLike = {
  viewport: PointerViewport;
  contentRect: PointerRect | null;
  interactiveRects?: PointerRect[] | null;
} | null;

type PointerInteractionWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  getBounds: () => PointerRect;
};

export type LinuxOverlayPointerInteractionDeps = {
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => PointerInteractionWindow | null;
  getCursorScreenPoint: () => PointerPoint;
  getSubtitleMeasurement: () => OverlayContentMeasurementLike;
  getRendererInteractiveHint: () => boolean;
  /** True when a modal/stats overlay owns input — leave interaction state to that logic. */
  shouldSuspend: () => boolean;
  getInteractionActive: () => boolean;
  setInteractionActive: (active: boolean) => void;
};

export const LINUX_OVERLAY_POINTER_POLL_INTERVAL_MS = 60;
// Padding (in window px) so the cursor doesn't have to land pixel-perfectly on the text.
const SUBTITLE_HIT_PADDING_PX = 6;

let pointerInteractionInterval: ReturnType<typeof setInterval> | null = null;

export function mapOverlayMeasurementForPointerInteraction(
  measurement: OverlayContentMeasurement | null,
): OverlayContentMeasurementLike {
  if (!measurement) return null;
  return {
    viewport: measurement.viewport,
    contentRect: measurement.contentRect,
    ...(measurement.interactiveRects ? { interactiveRects: measurement.interactiveRects } : {}),
  };
}

function isCursorOverRect(
  cursor: PointerPoint,
  bounds: PointerRect,
  viewport: PointerViewport,
  rect: PointerRect,
): boolean {
  if (!(bounds.width > 0) || !(bounds.height > 0)) return false;

  const scaleX = bounds.width / viewport.width;
  const scaleY = bounds.height / viewport.height;
  const left = bounds.x + rect.x * scaleX - SUBTITLE_HIT_PADDING_PX;
  const top = bounds.y + rect.y * scaleY - SUBTITLE_HIT_PADDING_PX;
  const right = left + rect.width * scaleX + SUBTITLE_HIT_PADDING_PX * 2;
  const bottom = top + rect.height * scaleY + SUBTITLE_HIT_PADDING_PX * 2;

  return cursor.x >= left && cursor.x <= right && cursor.y >= top && cursor.y <= bottom;
}

/** Hit-test the global cursor against subtitle bar rects, mapping viewport px → screen px. */
export function isCursorOverSubtitle(
  cursor: PointerPoint,
  bounds: PointerRect,
  measurement: OverlayContentMeasurementLike,
): boolean {
  if (!measurement) return false;
  const { viewport } = measurement;
  if (!(viewport.width > 0) || !(viewport.height > 0)) return false;

  const rects =
    Array.isArray(measurement.interactiveRects) && measurement.interactiveRects.length > 0
      ? measurement.interactiveRects
      : measurement.contentRect
        ? [measurement.contentRect]
      : [];

  return rects.some((rect) => isCursorOverRect(cursor, bounds, viewport, rect));
}

/**
 * Returns the desired interactive state, or null when the loop should not touch it
 * (overlay hidden/destroyed or another surface owns input).
 */
export function resolveDesiredOverlayInteractive(
  deps: LinuxOverlayPointerInteractionDeps,
): boolean | null {
  if (!deps.getVisibleOverlayVisible()) return false;
  if (deps.shouldSuspend()) return null;

  const mainWindow = deps.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    return null;
  }

  if (deps.getRendererInteractiveHint()) return true;
  return isCursorOverSubtitle(
    deps.getCursorScreenPoint(),
    mainWindow.getBounds(),
    deps.getSubtitleMeasurement(),
  );
}

export function tickLinuxOverlayPointerInteraction(deps: LinuxOverlayPointerInteractionDeps): void {
  const desired = resolveDesiredOverlayInteractive(deps);
  if (desired === null) return;
  if (deps.getInteractionActive() === desired) return;
  deps.setInteractionActive(desired);
}

export function ensureLinuxOverlayPointerInteractionLoop(
  deps: LinuxOverlayPointerInteractionDeps,
  platform: NodeJS.Platform = process.platform,
): void {
  if (pointerInteractionInterval !== null) return;
  if (platform !== 'linux') return;

  pointerInteractionInterval = setInterval(() => {
    tickLinuxOverlayPointerInteraction(deps);
  }, LINUX_OVERLAY_POINTER_POLL_INTERVAL_MS);
  pointerInteractionInterval.unref?.();
}

export function stopLinuxOverlayPointerInteractionLoop(): void {
  if (pointerInteractionInterval === null) return;
  clearInterval(pointerInteractionInterval);
  pointerInteractionInterval = null;
}
